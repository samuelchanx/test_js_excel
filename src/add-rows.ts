const Excel = require("exceljs");

const templateFilePath = "inputs/Credit summary NB V.1.2_20241025 template.xlsm";


async function readWorkbook(filePath: string) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filePath);
  const firstWorksheet = workbook.worksheets[1];
  const columnB = firstWorksheet.getColumn("C");
  columnB.eachCell(function (cell, rowNumber) {
    if (cell == "Insert Credit Commentary Here") {
      cell.value = "testing";
    }
  });

  return firstWorksheet;
}

readWorkbook(templateFilePath);