mod excel;

use std::{fs::{File}, io::{BufReader, BufRead, Write, BufWriter}, fmt::Display};
use neon::{prelude::*, types::Deferred};
// use simple_xlsx_writer::{WorkBook, Row as XLSRow, Cell};
use excel::{WorkBook, Row as XLSRow, Cell};

const BUFFER_CAPACITY: usize = 1_000_000;

struct Row<'a> (pub Vec<&'a str>);

impl<'a> Display for Row<'a> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let comb_str = self.0.join(",");
        write!(f, "{}", comb_str)
    }
}

fn read_file_liner<F>(filepath: String, fn_operation: &mut F) -> Result<(), std::io::Error>
    where
    F: FnMut(Row) -> Result<(), std::io::Error> {
    let mut file = File::open(filepath)?;
    let reader = BufReader::new(&file);

    for line in reader.lines() {
        let res_line = line.unwrap();
        let cols = res_line.split(",").collect::<Vec<&str>>();
        let row = Row(cols);
        fn_operation(row)?;
    }

    file.flush().expect("unable to close csv file");

    Ok(())
}

fn csv_to_excel(mut cx: FunctionContext) -> JsResult<JsPromise> {
    let csv_path = cx.argument::<JsString>(0)?.value(&mut cx);
    let xls_path = cx.argument::<JsString>(1)?.value(&mut cx);

    let channel = cx.channel();
    let (defer, promise) = cx.promise();

    std::thread::spawn(move || {
        write_xlsx(csv_path, xls_path, channel, defer);
    });

    Ok(promise)
}

fn write_xlsx(csv_path: String, xls_path: String, channel: Channel, defer: Deferred) {
    // let mut xls_file = BufWriter::with_capacity(BUFFER_CAPACITY, File::create(xls_path).unwrap());
    let mut xls_file = File::create(xls_path).unwrap();
    let mut workbook = Box::new(WorkBook::new(&mut xls_file ))
        .expect("unable to initiate excel workbook");
    let worksheet = workbook.get_new_sheet();
    
    let write_result = worksheet.write_sheet(true, |writer| {
        let mut operation = |row: Row| {
            let mut xls_row = XLSRow::new();
    
            for col in row.0 {
                xls_row.add_cell(Cell::from(col));
            }
    
            writer.write_row(xls_row)
        };
        read_file_liner(csv_path, &mut operation)
    });
    
    workbook.finish().unwrap();
    xls_file.flush().unwrap();

    // let res = read_file_liner(csv_path, &mut operation);
    defer.settle_with(&channel, move |mut cx| {
        match write_result {
            Ok(()) => Ok(cx.boolean(true)),
            Err(err) => cx.throw_error(err.to_string()),
        }
    });
}

#[neon::main]
fn main(mut cx: ModuleContext) -> NeonResult<()> {
    cx.export_function("CsvToExcel", csv_to_excel)?;
    Ok(())
}
