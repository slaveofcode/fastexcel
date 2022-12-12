# FastExcel

> This project need Rust to be installed, check here for [Rust installation instruction](https://www.rust-lang.org/tools/install)

> This project using [Rust](https://www.rust-lang.org) and [Neon](https://neon-bindings.com) as a binding to Rust to execute fast and efficient memory usage for generating XLSX document from NodeJs. 

> This project cannot be executed via NVM based NodeJs, you should deactivate (via `nvm deactivate`) and use a normal version installation of NodeJs.

Writing a large amount of data into Excel file is not a trivial task when you have a limited memory (RAM) allocated. Especially when working at a small node on the server. This library is created to solve that problem, using the efficiency of Rust while generating XLSX from CSV.

### Installation

    npm i -D cargo-cp-artifact

    npm i fastexcel

### How it works

1. Generate the CSV
2. Convert the CSV to XLSX

The CSV generation is happen on the NodeJs side, and converting XLSX file is on Rust side (via Neon)

### Example Usage

```js
// dummy-excel.js
const path = require('path');
const { CsvFileWriter, Converter } = require("fastexcel");

const main = async () => {
  const src = path.join(process.cwd(), 'example/source.csv');
  const dst = path.join(process.cwd(), 'example/generated.xlsx');

  console.log('src', src);
  console.log('dst', dst);

  const cols = [];
  const totalCols = 200; // 200 columns
  for (let i = 0; i < totalCols; i++) {
    cols.push('Col ' + (i+1));
  }

  const writer = new CsvFileWriter(src, cols);

  const totalRows = 1_000_000; // 1 million rows
  for (let i = 0; i < totalRows; i++) {
    let row = [];
    
    for (let i = 0; i < totalCols; i++) {
      row.push('Col No ' + (i+1));
    }

    row.push(row);
    await writer.write(row);
  }

  await writer.close();

  // Part 2: Convert csv to excel
  const res = await Converter.toXLSX(
    src,
    dst
  );
};

main();
```
