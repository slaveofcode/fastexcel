# FastExcel

### Example

```
// Part 1: Put Data to CSV/Text
const writer = new CsvFileWriter("./test/source-lib.csv", [
  "No",
  "Name",
  "Gender",
]);

const rows = [
  [1, "John", "Male"],
  [2, "Doe", "Male"],
];

for (const row of rows) {
  await writer.write(row);
}

await writer.close();

// Part 2: Convert csv to excel
const res = await Converter.toXLSX(
  "./test/source-lib.csv",
  "./test/result-lib.xlsx",
);
```
