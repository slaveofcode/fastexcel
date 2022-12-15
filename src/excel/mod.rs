//! # simple_xlsx_writer
//! This is a very simple XLSX writer library.
//!
//! This is not feature rich and it is not supposed to be. A lot of the design was based on the work of [simple_excel_writer](https://docs.rs/simple_excel_writer/latest/simple_excel_writer/) and I recomend you to check that crate.
//!
//! The main idea of this create is to help you build XLSX files using very little RAM.
//! I created it to use in my web assembly site [csv2xlsx](https://csv2xlsx.com).
//!
//! Basically, you just need to pass an output that implements [Write](std::io::Write) and [Sink](std::io::Sink) to the [WorkBook](crate::WorkBook). And while you are writing the file, it wil be written directly to the output already compressed. So, you could stream directly into a file using very little RAM. Or even write to the memory and still not use that much memory because the file will be already compressed.
//!
//! ## Example
//! ```rust
//! use simple_xlsx_writer::{row, Row, WorkBook};
//! use std::fs::File;
//! use std::io::Write;
//!
//! fn main() -> std::io::Result<()> {
//!     let mut files = File::create("example.xlsx")?;
//!     let mut workbook = WorkBook::new(&mut files)?;
//!     let header_style = workbook.create_cell_style((255, 255, 255), (0, 0, 0));
//!     workbook.get_new_sheet().write_sheet(|sheet_writer| {
//!         sheet_writer.write_row(row![("My", &header_style), ("Sample", &header_style), ("Header", &header_style)])?;
//!         sheet_writer.write_row(row![1, 2, 3])?;
//!         Ok(())
//!     })?;
//!     workbook.get_new_sheet().write_sheet(|sheet_writer| {
//!         sheet_writer.write_row(row![("Another", &header_style), ("Sheet", &header_style), ("Header", &header_style)])?;
//!         sheet_writer.write_row(row![1.32, 2.43, 3.54])?;
//!         Ok(())
//!     })?;
//!     workbook.finish()?;
//!     files.flush()?;
//!     Ok(())
//! }
//! ```
mod row;
mod sheet;
mod sheet_writer;
mod workbook;

pub use row::{Cell, CellValue, Row};
pub use sheet::{Sheet};
pub use sheet_writer::{SheetWriter};
pub use workbook::{CellStyle, WorkBook};

#[macro_export]
macro_rules! row {
    ($( $x:expr ),*) => {
        {
            let mut row = Row::new();
            $(row.add_cell($x.into());)*
            row
        }
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{open_workbook_from_rs, Reader, Xlsx};
    use std::io::{Cursor, Result as IoResult};

    // Very simple smoke test.
    #[test]
    fn it_works() -> IoResult<()> {
        let mut cursor = Cursor::new(Vec::new());
        let mut workbook = WorkBook::new(&mut cursor)?;
        let cell_style = workbook.create_cell_style((255, 255, 255), (0, 0, 0));
        let sheet_1 = workbook.get_new_sheet();
        sheet_1.write_sheet(false, |sheet_writer| {
            sheet_writer.write_row(row!(
                (1, &cell_style),
                (10.3, &cell_style),
                (54.3, &cell_style)
            ))?;
            sheet_writer.write_row(row!("ola", "text", "tree"))?;
            sheet_writer.write_row(row!(true, false, false, false))?;
            Ok(())
        })?;
        let sheet_2 = workbook.get_new_sheet();
        sheet_2.write_sheet(false, |sheet_writer| {
            sheet_writer.write_row(row!(1, 2, 3, 4, 4))?;
            sheet_writer.write_row(row!("one", "two", "three"))?;
            sheet_writer.write_row(row!("Another row"))?;
            Ok(())
        })?;
        workbook.finish()?;
        Ok(())
    }

    #[test]
    fn test_is_valid_report() -> IoResult<()> {
        let mut cursor = Cursor::new(Vec::new());
        let mut workbook = WorkBook::new(&mut cursor)?;
        let sheet_1 = workbook.get_new_sheet();
        sheet_1.write_sheet(false, |sheet_writer| {
            sheet_writer.write_row(row!(1, 10.3, 54.3))?;
            sheet_writer.write_row(row!("ola", "text", "tree"))?;
            sheet_writer.write_row(row!(true, false, false, false))?;
            Ok(())
        })?;
        let sheet_2 = workbook.get_new_sheet();
        sheet_2.write_sheet(false, |sheet_writer| {
            sheet_writer.write_row(row!(1, 2, 3, 4, 4))?;
            sheet_writer.write_row(row!("one", "two", "three"))?;
            sheet_writer.write_row(row!("Another row"))?;
            Ok(())
        })?;
        workbook.finish()?;
        let result = xlsx_to_vec(cursor);
        assert_eq!(
            vec![
                vec![
                    vec!["1", "10.3", "54.3", ""],
                    vec!["ola", "text", "tree", ""],
                    vec!["true", "false", "false", "false"]
                ],
                vec![
                    vec!["1", "2", "3", "4", "4"],
                    vec!["one", "two", "three", "", ""],
                    vec!["Another row", "", "", "", ""],
                ]
            ],
            result
        );
        Ok(())
    }

    fn xlsx_to_vec(cursor: Cursor<Vec<u8>>) -> Vec<Vec<Vec<String>>> {
        let mut xlsx_reader: Xlsx<_> = open_workbook_from_rs(cursor).unwrap();
        let mut result = (1..5)
            .into_iter()
            .map(
                |sheet_n| match xlsx_reader.worksheet_range(&format!("Sheet {}", sheet_n)) {
                    Some(result_calamine) => result_calamine
                        .unwrap()
                        .rows()
                        .map(|row| row.into_iter().map(|column| column.to_string()).collect())
                        .collect::<Vec<Vec<String>>>(),
                    None => Vec::new(),
                },
            )
            .collect::<Vec<Vec<Vec<String>>>>();

        result.reverse();
        result = result
            .into_iter()
            .skip_while(|row| row.is_empty())
            .collect::<Vec<Vec<Vec<String>>>>();
        result.reverse();
        result
    }
}
