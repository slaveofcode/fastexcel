use std::{fs::File, error::Error, io::{Read, BufReader, BufRead, Write}, fmt::Display};
use neon::{prelude::*};
use simple_xlsx_writer::{WorkBook, Row as XLSRow, Cell};

struct Row<'a> (pub Vec<&'a str>);

impl<'a> Display for Row<'a> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let comb_str = self.0.join(",");
        write!(f, "{}", comb_str)
    }
}

#[allow(dead_code)]
fn read_file_buffer(filepath: String) -> Result<(), Box<dyn Error>> {
    const BUFFER_LEN: usize = 512;
    let mut buffer = [0u8; BUFFER_LEN];
    let mut file = File::open(filepath)?;

    loop {
        let read_count = file.read(&mut buffer)?;
        let buff = &buffer[..read_count];
        let line = String::from_utf8_lossy(buff);
        println!("line: {}", line);

        if read_count != BUFFER_LEN {
            break;
        }
    }
    Ok(())
}

fn read_file_liner<F>(filepath: String, fn_operation: &mut F) -> Result<(), std::io::Error>
    where
    F: FnMut(&mut Row) -> Result<(), std::io::Error> {
    let file = File::open(filepath)?;
    let reader = BufReader::new(file);

    for line in reader.lines() {
        let res_line = line.unwrap();
        let mut row = Row(vec![&res_line]);
        fn_operation(&mut row)?;
    }

    Ok(())
}

fn csv_to_excel(mut cx: FunctionContext) -> JsResult<JsPromise> {
    let csv_path = cx.argument::<JsString>(0)?.value(&mut cx);
    let xls_path = cx.argument::<JsString>(1)?.value(&mut cx);

    let channel = cx.channel();
    let (defer, promise) = cx.promise();

    defer.settle_with(&channel, move |mut cx| {
        let mut xls_file = File::create(&xls_path)
            .expect(&format!("unable to create xlsx file: {}", &xls_path));
        let mut workbook = WorkBook::new(&mut xls_file )
            .expect("unable to initiate excel workbook");

        let res = workbook.get_new_sheet().write_sheet(|writer| {
            let mut operation = |row: &mut Row| {
                let mut xls_row = XLSRow::new();

                for col in &row.0 {
                    xls_row.add_cell(Cell::from(*col));
                }

                writer.write_row(xls_row)
            };
            read_file_liner(csv_path, &mut operation)
        });

        workbook.finish().unwrap();
        xls_file.flush().unwrap();
        
        match res {
            Ok(()) => Ok(cx.boolean(true)),
            Err(err) => cx.throw_error(err.to_string()),
        }
    });

    Ok(promise)
}

#[neon::main]
fn main(mut cx: ModuleContext) -> NeonResult<()> {
    cx.export_function("CsvToExcel", csv_to_excel)?;
    Ok(())
}
