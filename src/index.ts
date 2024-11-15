// import readXlsxFile from "read-excel-file/node";
import * as ExcelJS from "exceljs";
import * as xlsx from "xlsx";

async function main() {
  console.log("test start now");
  try {
    // xlsxPackageTest()
    await test3();
    // await execljsPackageTest();
  } catch (e) {
    console.error(e);
  }
}

async function test3() {
  const Excel = require("exceljs");

  // target file
  const wb = new Excel.Workbook();

  // read from a file
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(
    "src/Credit summary NB V.1.2_20200504 SA EU WU.xlsm"
  );
  const ws = workbook.worksheets[0];
  console.log(ws.rowBreaks);
  console.log(workbook.worksheets[1].rowBreaks);
  console.log(workbook.worksheets[2].rowBreaks);
  console.log(ws.pageSetup.printArea);

  for (const worksheet of workbook.worksheets) {
    // console.log(JSON.stringify(worksheet.pageSetup));
    const ws = wb.addWorksheet(worksheet.name, {
      pageSetup: worksheet.pageSetup,
      properties: worksheet.properties,
      headerFooter: worksheet.headerFooter,
      views: worksheet.views,
    });
    const lastRowNum = worksheet.rowCount;

    worksheet.eachRow((row, rowNum) => {
      if (rowNum !== lastRowNum) {
        const r = ws.addRow();
        Object.assign(r, row);
        r.hidden = row.hidden;

        // r.height = row.height;
        // r.outlineLevel = row.outlineLevel;
      }
    });
    for (let i = 1; i <= worksheet.columnCount; i++) {
      ws.getColumn(i).width = worksheet.getColumn(i).width;
    }
    for (let i = 1; i <= worksheet.rowCount; i++) {
      if (worksheet.getRow(i).hidden) {
        console.log("Hidden row", i);
      }
      const row = ws.getRow(i-1)
      row.hidden = worksheet.getRow(i).hidden;
      row.commit()
    }
    ws.pageSetup.printArea = worksheet.pageSetup.printArea;
    if (worksheet.name === "Project Summary") {
      //   console.log(worksheet.pageSetup.printArea);
      //   console.log(ws.pageSetup.printArea);
    }
    
    ws.properties = worksheet.properties;
    ws.headerFooter = worksheet.headerFooter;
    ws.views = worksheet.views;
  }

  const sheet = wb.worksheets[0];
  const cell = sheet.getCell("A1");
  cell.value = "test";
  console.log(cell);
  cell.style = {
    ...cell.style,
    font: {
      bold: true,
      color: {
        argb: "FFFFFFFF",
      },
    },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        // argb: 'FFFFFFFF'
        argb: "00FF0000",
      },
    },
  };

  wb.xlsx.writeFile("33-copy.xlsx").then(() => {
    console.log("done!");
  });
}

/// exceljs method call
async function execljsPackageTest() {
  const work = new ExcelJS.Workbook();
  const work2 = new ExcelJS.Workbook();
  const newWork = new ExcelJS.Workbook();
  const book1 = await work.xlsx.readFile("src/Book1.xlsx");
  const book2 = await work2.xlsx.readFile("src/Book2.xlsx");
  // console.log(book1.worksheets[0])
  const allBook = [...book1.worksheets, ...book2.worksheets];
  for (const b of allBook) {
    const emptyBook = newWork.addWorksheet(b.name);
    emptyBook.model = b.model;
  }
  const allSheet = newWork.worksheets.map((e) => e.name);
  const sheet = newWork.getWorksheet(allSheet[0]);
  const cell = sheet.getCell("A1");
  console.log(cell);
  cell.style = {
    ...cell.style,
    font: {
      bold: true,
    },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: {
        // argb: 'FFFFFFFF'
        argb: "00FF0000",
      },
    },
  };
  await newWork.xlsx.writeFile("test2.xlsx");
}

/// xlsx package method call
function xlsxPackageTest() {
  const book1 = xlsx.readFile(
    "src/Credit summary NB V.1.2_20200504 SA EU WU.xlsm"
  );
  const book2 = xlsx.readFile("src/Book2.xlsx");
  const newBook = xlsx.utils.book_new();
  book1.SheetNames.forEach((e) => {
    xlsx.utils.book_append_sheet(newBook, book1.Sheets[e]);
  });
  // xlsx.utils.book_append_sheet(newBook,book1.Sheets['main'])
  // xlsx.utils.book_append_sheet(newBook,book1.Sheets['sub'], "YO2")
  // xlsx.utils.book_append_sheet(newBook,book2.Sheets['Jan'], "YO3")
  // xlsx.utils.book_append_sheet(newBook,book2.Sheets['Feb'], "YO4")
  // console.log(newBook);
  const row = xlsx.utils.sheet_add_aoa(newBook.Sheets["YO"], [["Column1"]], {
    origin: "A1",
  });
  console.log(row);
  xlsx.writeFile(newBook, "test.xlsx");
}

main();
