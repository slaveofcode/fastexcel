# FastExcel

> This project using [Neon](https://neon-bindings.com) as a binding to Rust to execute fast and efficient memory usage for generating XLSX document. 

### Installation

    npm i fastexcel

### Example

```
// index.js
const path = require('path');
const { CsvFileWriter, Converter } = require("fastexcel");

const main = async () => {
  const src = path.join(process.cwd(), 'example/source.csv');
  const dst = path.join(process.cwd(), 'example/generated.xlsx');

  console.log('src', src);
  console.log('dst', dst);

  const cols = [];
  const totalCols = 200;
  for (let i = 0; i < totalCols; i++) {
    cols.push('Col ' + (i+1));
  }

  const writer = new CsvFileWriter(src, cols);

  const totalRows = 1000000; // 1jt rows
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
