import { commentedWorksheets } from "./constants";

const Excel = require("exceljs");

const templateFilePath =
  "inputs/Credit summary NB V.1.2_20241025 template.xlsm";

async function readWorkbook(filePath: string) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filePath);

  async function updateOneWorksheet(worksheet) {
    console.log("Processing worksheet:", worksheet.name);
    const columnB = worksheet.getColumn("C");

    let lastProcessedRowNumber = -1;
    columnB.eachCell(function (cell, rowNumber) {
      if (cell) {
        if (cell.value == "Insert Credit Commentary Here") {
          const cellOfConstested = worksheet.getCell(`C${rowNumber + 1}`);
          cellOfConstested.value = "Insert Contested Reason Here";

          const cellOfDiscussion = worksheet.getCell(`C${rowNumber + 2}`);
          cellOfDiscussion.value = "Insert Discussion Here";

          // const rowOfDiscussion = [];
          // rowOfDiscussion[3] = "Insert Discussion Here";
          // worksheet.insertRow(rowNumber + 2, rowOfDiscussion, "i");
        }
      }
      // workbook;
    });
  }

  for (const worksheet of commentedWorksheets) {
    await updateOneWorksheet(workbook.getWorksheet(worksheet));
  }

  await workbook.xlsx.writeFile("outputs/test.xlsx");
  console.log("done!");
}

readWorkbook(templateFilePath);
