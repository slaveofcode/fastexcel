const { unlinkSync, statSync } = require("fs");
const { convertCsvToExcel } = require("../dist/tool");

jest.setTimeout(180000);

beforeAll(() => {
  const xlsResult = ["100k", "300k", "500k", "1m"];
  for (const pattern of xlsResult) {
    try {
      const path = `./test/result-${pattern}.xlsx`;
      if (statSync(path)) {
        unlinkSync(path);
      }
    } catch (err) {}
  }
});

test("Must be able to convert 100K rows to Excel", async () => {
  const res = await convertCsvToExcel(
    "./test/100k.csv",
    "./test/result-100.xlsx"
  );

  expect(res).toEqual(true);
});

test("Must be able to convert 300K rows to Excel", async () => {
  const res = await convertCsvToExcel(
    "./test/300k.csv",
    "./test/result-300k.xlsx"
  );

  expect(res).toEqual(true);
});

