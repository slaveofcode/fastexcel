const fs = require('fs');
const { CsvFileWriter, Converter } = require("../dist");

jest.setTimeout(240000);

beforeEach(() => {
  try {
    const path = `./test/source-lib.csv`;
    if (statSync(path)) {
      unlinkSync(path);
    }
  } catch (err) {}
});

test("Library is able to create csv", async () => {
  const writer = new CsvFileWriter('./test/source-lib.csv', [
    "No",
    "Name",
    "Gender",
  ]);

  // const rows = [
  //   [1, "John", "Male"],
  //   [2, "Doe", "Male"],
  // ];

  // for (const row of rows) {
  //   await writer.write(row);
  // }

  for (let row = 1; row <= 470_000; row++) {
    const csvRow = [];
    for (let col = 1; col <= 200; col++) {
      csvRow.push(`col no ${col}`);
    }
    await writer.write(csvRow);
  }

  await writer.close();

  

//   const content = fs.readFileSync('./test/source-lib.csv', 'utf8');

//   expect(content).toEqual(`No,Name,Gender
// 1,John,Male
// 2,Doe,Male
// `)
});

test("Library is able to create and convert xlsx", async () => {
  // const writer = new CsvFileWriter("./test/source-lib.csv", [
  //   "No",
  //   "Name",
  //   "Gender",
  // ]);

  // const rows = [
  //   [1, "John", "Male"],
  //   [2, "Doe", "Male"],
  // ];

  // for (const row of rows) {
  //   await writer.write(row);
  // }

  // await writer.close();

  const s = setTimeout(() => {
    console.log('event loop running...')
  }, 30000)

  const res = await Converter.toXLSX("./test/source-lib.csv", "./test/result-lib.xlsx");

  clearTimeout(s);

  expect(res).toEqual(true);
});
