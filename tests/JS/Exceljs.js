import { equal } from 'assert';
import { Xlsx } from './FsSpreadsheet.Exceljs/Xlsx.fs.js';
import { Excel } from './FsSpreadsheet.Exceljs/fable_modules/Fable.Exceljs.1.3.6/ExcelJs.fs.js';
import { toJsWorkbook, toFsWorkbook } from "./FsSpreadsheet.Exceljs/Workbook.fs.js";

describe('FsSpreadsheet.Exceljs', function () {
    describe('read', function () {
        it('values', async () => {
            //const wb = await Xlsx.fromXlsxFile(path)
            const path = "tests/JS/TestFiles/ReadTable.xlsx"; // path always from package.json 
            const wb = new Excel.Workbook();
            await wb.xlsx.readFile(path)
            const fswb = toFsWorkbook(wb)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
            let ws = worksheets[0]
            equal(ws.name, "Tabelle1")
            // row 1
            equal(ws.GetCellAt(5, 4).Value, "My number column")
            equal(ws.GetCellAt(5, 5).Value, "My boolean column")
            equal(ws.GetCellAt(5, 6).Value, "My string column")
            // row 2
            equal(ws.GetCellAt(6, 4).Value, 2)
            equal(ws.GetCellAt(6, 5).ValueAsBool(), true)
            equal(ws.GetCellAt(6, 6).Value, "row2")
            // row 3
            equal(ws.GetCellAt(7, 4).Value, 20)
            equal(ws.GetCellAt(7, 5).ValueAsBool(), false)
            equal(ws.GetCellAt(7, 6).Value, "row20")
        });
        it('table', async () => {
            //const wb = await Xlsx.fromXlsxFile(path)
            const path = "tests/JS/TestFiles/ReadTable.xlsx";
            const wb = new Excel.Workbook();
            await wb.xlsx.readFile(path)
            const fswb = toFsWorkbook(wb)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
            let ws = worksheets[0]
            let table = ws.Tables
            equal(ws.name, "Tabelle1")
        });
        it('combined function', async () => {
            const path = "tests/JS/TestFiles/ReadTable.xlsx";
            console.log("start")
            const wb = await Xlsx.fromXlsxFile(path)
            console.log(wb)
        });
    });
});