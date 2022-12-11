const path = require('path');
const { CsvWriter } = require('./csv-writer');
const { convertCsvToExcel } = require('./tool');

class CsvFileWriter {
  constructor(csvFilePath, columns) {
    this.csvWriter = new CsvWriter(csvFilePath, columns);
  }

  async write(row) {
    await this.csvWriter.write(row);
  }

  close() {
    return this.csvWriter.pack();
  }
}

class Converter {
  static async toXLSX(srcCsv, xlsFilePath) {
    return await convertCsvToExcel(srcCsv, xlsFilePath);
  }
}

module.exports = {
  CsvFileWriter,
  Converter,
};
