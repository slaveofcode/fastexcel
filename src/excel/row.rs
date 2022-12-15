use crate::excel::CellStyle;
use std::io::{Result as IoResult, Write};

/// A row of a sheet. You can also create it using the macro `row!`
#[derive(Clone, Debug)]
pub struct Row<'a> {
    cells: Vec<Cell<'a>>,
}

/// A Cell of a row. It has a [CellValue](CellValue) and an optional [CellStyle](CellStyle)
#[derive(Clone, Debug)]
pub struct Cell<'a> {
    value: CellValue,
    style: Option<&'a CellStyle>,
}

impl<'a> Row<'a> {
    pub fn new() -> Self {
        Self { cells: Vec::new() }
    }

    pub fn add_cell(&mut self, cell: Cell<'a>) {
        self.cells.push(cell);
    }

    pub fn cells(self) -> Vec<Cell<'a>> {
        self.cells
    }
}

/// A cell value. Right now, we can represent bool, f64 and strings
#[derive(Clone, Debug)]
pub enum CellValue {
    Bool(bool),
    Number(f64),
    String(String),
}

impl<'a> Cell<'a> {
    pub fn write(
        &self,
        column_index: u8,
        row_index: usize,
        writer: &mut impl Write,
    ) -> IoResult<()> {
        let ref_id = ref_id(column_index, row_index);
        match &self.value {
            CellValue::Bool(b) => write!(
                writer,
                "<c r=\"{}\" t=\"b\"{}><v>{}</v></c>\n",
                ref_id,
                self.cell(),
                if *b { 1 } else { 0 }
            ),
            CellValue::Number(number) => {
                write!(
                    writer,
                    "<c r=\"{}\"{}><v>{}</v></c>\n",
                    ref_id,
                    self.cell(),
                    number
                )
            }
            CellValue::String(string) => {
                write!(
                    writer,
                    "<c r=\"{}\" t=\"str\"{}><v>{}</v></c>\n",
                    ref_id,
                    self.cell(),
                    escape_xml(string.as_str())
                )
            }
        }
    }

    fn cell(&self) -> String {
        match self.style {
            None => "".to_string(),
            Some(style) => format!(" s=\"{}\"", style.get_id()),
        }
    }
}

pub fn escape_xml(str: &str) -> String {
    let mut result = String::new();
    for c in str.chars() {
        match c {
            '&' => result.push_str("&amp;"),
            '<' => result.push_str("&lt;"),
            '>' => result.push_str("&gt;"),
            '\'' => result.push_str("&apos;"),
            '\"' => result.push_str("&quot;"),
            _ => result.push(c),
        }
    }
    result
}

fn ref_id(column_index: u8, row_index: usize) -> String {
    format!("{}{}", column_letter(column_index), row_index)
}

fn column_letter(column_index: u8) -> String {
    let mut result = Vec::new();
    let mut column_index = column_index as i16;
    while column_index >= 0 {
        result.push(number_to_letter((column_index % 26) as u8));
        column_index = column_index / 26 - 1;
    }
    result.reverse();
    result.into_iter().collect()
}

fn number_to_letter(number: u8) -> char {
    (b'A' + number) as char
}

impl<'a, T> From<T> for Cell<'a>
where
    T: Into<CellValue>,
{
    fn from(value: T) -> Self {
        Cell {
            value: value.into(),
            style: None,
        }
    }
}

impl<'a, T> From<(T, &'a CellStyle)> for Cell<'a>
where
    T: Into<CellValue>,
{
    fn from(value: (T, &'a CellStyle)) -> Self {
        Cell {
            value: value.0.into(),
            style: Some(value.1),
        }
    }
}

impl From<bool> for CellValue {
    fn from(data: bool) -> Self {
        Self::Bool(data)
    }
}

impl From<f64> for CellValue {
    fn from(data: f64) -> Self {
        Self::Number(data)
    }
}

impl From<i32> for CellValue {
    fn from(data: i32) -> Self {
        Self::Number(data.into())
    }
}

impl From<String> for CellValue {
    fn from(data: String) -> Self {
        Self::String(data)
    }
}

impl From<&str> for CellValue {
    fn from(data: &str) -> Self {
        Self::String(data.to_string())
    }
}

impl<'a, T> From<Vec<T>> for Row<'a>
where
    T: Into<Cell<'a>>,
{
    fn from(vec: Vec<T>) -> Row<'a> {
        let mut row = Row::new();
        for value in vec.into_iter() {
            row.add_cell(value.into());
        }
        row
    }
}
