use crate::excel::Row;
use std::io::{Result as IoResult, Write};

pub struct SheetWriter<W>
where
    W: Write,
{
    writer: W,
    row_index: usize,
    written_footer: bool,
}

impl<W> SheetWriter<W>
where
    W: Write,
{
    /// Writes a row into the sheet.
    pub fn write_row(&mut self, row: Row) -> IoResult<()> {
        self.row_index += 1;
        write!(self.writer, "<row r=\"{}\">\n", self.row_index)?;
        for (i, c) in row.cells().into_iter().enumerate() {
            c.write(i as u8, self.row_index, &mut self.writer)?;
        }
        write!(self.writer, "\n</row>\n")?;
        self.writer.flush()
        // Ok(())
    }

    /// Finish the sheet. Necessary to be called if you got the [SheetWriter](SheetWriter) from [Sheet::sheet_writer](Sheet::sheet_writer). We also try to execute this in the [Drop](SheetWriter::drop), but it is a good practice to always finish the sheet.
    pub fn finish(mut self) -> IoResult<()> {
        self.write_footer()
    }

    pub fn start(writer: W) -> IoResult<Self>
    where
        W: Write,
    {
        let mut writer = Self {
            writer,
            row_index: 0,
            written_footer: false,
        };
        writer.write_header()?;
        Ok(writer)
        
    }

    fn write_header(&mut self) -> IoResult<()> {
        write!(
            self.writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#
        )?;
        write!(self.writer, "\n<sheetData>\n")?;
        self.writer.flush()
        // Ok(())
    }

    fn write_footer(&mut self) -> IoResult<()> {
        self.written_footer = true;
        write!(self.writer, "\n</sheetData>\n</worksheet>\n").expect("unable write sheet footer");
        self.writer.flush()
    }
}

impl<W> Drop for SheetWriter<W>
where
    W: Write,
{
    /// Drops the [SheetWriter](SheetWriter) and tries to finish it if not already finished. This might panic if we fail to write the footer of the sheet.
    fn drop(&mut self) {
        if self.written_footer == false {
            self.write_footer().expect("Error written sheet footer");
        }
    }
}