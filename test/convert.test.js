const { unlinkSync, statSync } = require("fs");
const { convertCsvToExcel } = require("../dist/tool");

jest.setTimeout(60000);

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
  const fnNoThrow = async () => {
    return await convertCsvToExcel("./test/100k.csv", "./test/result-100.xlsx");
  };

  const res = await convertCsvToExcel(
    "./test/100k.csv",
    "./test/result-100.xlsx"
  );

  expect(fnNoThrow).not.toThrow();
  expect(res).toEqual(true);
});

test("Must be able to convert 300K rows to Excel", async () => {
  const fnNoThrow = async () => {
    return await convertCsvToExcel(
      "./test/300k.csv",
      "./test/result-300k.xlsx"
    );
  };

  const res = await convertCsvToExcel(
    "./test/300k.csv",
    "./test/result-300k.xlsx"
  );

  expect(fnNoThrow).not.toThrow();
  expect(res).toEqual(true);
});

test("Must be able to convert 500K rows to Excel", async () => {
  const fnNoThrow = async () => {
    return await convertCsvToExcel(
      "./test/500k.csv",
      "./test/result-500k.xlsx"
    );
  };

  const res = await convertCsvToExcel(
    "./test/500k.csv",
    "./test/result-500k.xlsx"
  );

  expect(fnNoThrow).not.toThrow();
  expect(res).toEqual(true);
});

test("Must be able to convert 1M rows to Excel", async () => {
  const fnNoThrow = async () => {
    return await convertCsvToExcel("./test/1m.csv", "./test/result-1m.xlsx");
  };

  const res = await convertCsvToExcel("./test/1m.csv", "./test/result-1m.xlsx");

  expect(fnNoThrow).not.toThrow();
  expect(res).toEqual(true);
});
