use std::io::{Result as IoResult, Seek, Write};
use zip::{write::FileOptions, ZipWriter};
use crate::excel::SheetWriter;

/// A XLSX sheet.
pub struct Sheet<'a, W>
where
    W: Write + Seek,
{
    id: usize,
    zip_writer: &'a mut ZipWriter<W>,
}

/// Responsible to write a sheet into the workbook.

impl<'a, W> Sheet<'a, W>
where
    W: Write + Seek,
{
    pub(crate) fn new(id: usize, zip_writer: &'a mut ZipWriter<W>) -> Self {
        Self { id, zip_writer }
    }

    /// Receives a closure that will write the sheet. The closure receive a [SheetWriter](SheetWriter) that can be used to write the rows into the sheet.
    /// You don't need to call [finish](SheetWriter::finish) as it will be called for you.
    pub fn write_sheet<T>(
        self,
        is_large: bool,
        function: impl FnOnce(&mut SheetWriter<&mut ZipWriter<W>>) -> IoResult<T>,
    ) -> IoResult<T> {
        let options = FileOptions::default().large_file(is_large);
        self.zip_writer
            .start_file(format!("xl/worksheets/sheet{}.xml", self.id), options)?;
        let mut sheet_writer = SheetWriter::start(&mut *self.zip_writer)?;
        let result = function(&mut sheet_writer)?;
        sheet_writer.finish()?;
        Ok(result)
    }

    /// Antoher way to write a sheet. Insted of using a closure that has access to the [SheetWriter](SheetWriter). This returns the [SheetWriter](SheetWriter) directly and you can use it to write the sheet.
    /// You need to call [finish](SheetWriter::finish).
    /// set is_large to true when the file is approximately bigger than 4GiB
    pub fn sheet_writer(self, is_large: bool) -> IoResult<SheetWriter<&'a mut ZipWriter<W>>> {
        let options = FileOptions::default().large_file(is_large);
        self.zip_writer
            .start_file(format!("xl/worksheets/sheet{}.xml", self.id), options)?;
        Ok(SheetWriter::start(&mut *self.zip_writer)?)
    }
}
