// import readXlsxFile from "read-excel-file/node";
import * as ExcelJS from 'exceljs';
import * as xlsx from 'xlsx';

async function main () {
    console.log('test start now')
    try{
        
        // xlsxPackageTest()
        await execljsPackageTest()
    } catch (e) {
        console.error(e)
    }
}

/// exceljs method call
async function execljsPackageTest() {
    const work = new ExcelJS.Workbook()
    const work2 = new ExcelJS.Workbook()
    const newWork = new ExcelJS.Workbook()
    const book1 = await work.xlsx.readFile('src/Book1.xlsx')
    const book2 = await work2.xlsx.readFile('src/Book2.xlsx')
    // console.log(book1.worksheets[0])
    const allBook = [...book1.worksheets , ...book2.worksheets]
    for (const b of allBook) {
        const emptyBook = newWork.addWorksheet(b.name)
        emptyBook.model = b.model;
    }
    const allSheet = newWork.worksheets.map((e) => e.name)
    const sheet = newWork.getWorksheet(allSheet[0])
    const cell = sheet.getCell('A1')
    console.log(cell)
    cell.style = {
        ...cell.style,
        font: {
            bold: true
        },
        fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {
                // argb: 'FFFFFFFF'
                argb: '00FF0000'
            }
        }
    }
    await newWork.xlsx.writeFile('test2.xlsx')
}


/// xlsx package method call
function xlsxPackageTest() {
    const book1 = xlsx.readFile('src/Book1.xlsx')
    const book2 = xlsx.readFile('src/Book2.xlsx')
    const newBook = xlsx.utils.book_new()
    xlsx.utils.book_append_sheet(newBook,book1.Sheets['main'], "YO")
    xlsx.utils.book_append_sheet(newBook,book1.Sheets['sub'], "YO2")
    xlsx.utils.book_append_sheet(newBook,book2.Sheets['Jan'], "YO3")
    xlsx.utils.book_append_sheet(newBook,book2.Sheets['Feb'], "YO4")
    // console.log(newBook);
    const row = xlsx.utils.sheet_add_aoa(newBook.Sheets['YO'], [['Column1']], {
        origin: 'A1'
    })
    console.log(row)
    xlsx.writeFile(newBook,'test.xlsx')

}

main()