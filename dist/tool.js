const lib = require("../index.node");

const convertCsvToExcel = (csvSrc, xlsDst) => {
  return lib.CsvToExcel(csvSrc, xlsDst);
};

module.exports = {
  convertCsvToExcel,
};
