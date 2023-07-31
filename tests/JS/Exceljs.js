import { equal } from 'assert';
import { Xlsx } from './FsSpreadsheet.Exceljs/Xlsx.fs.js';
import { Excel } from './FsSpreadsheet.Exceljs/fable_modules/Fable.Exceljs.1.3.6/ExcelJs.fs.js';
import { FsWorkbook } from "./FsSpreadsheet.Exceljs/FsSpreadsheet/FsWorkbook.fs.js";
import { FsRangeAddress_$ctor_Z721C83C5, FsRangeAddress__get_Range } from "./FsSpreadsheet.Exceljs/FsSpreadsheet/Ranges/FsRangeAddress.fs.js";
import { FsTable } from "./FsSpreadsheet.Exceljs/FsSpreadsheet/Tables/FsTable.fs.js";
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
            equal(ws.Name, "Tabelle1")
            let tables = ws.Tables
            equal(tables.length, 1)
            let table = ws.Tables[0]
            equal(table.Name, "Table1")
        });
    })
    describe('write', function () {
        it('roundabout', async () => {
            const path = "tests/JS/TestFiles/WriteTable.xlsx"
            const path2 = "tests/JS/TestFiles/WriteTable2.xlsx"
            const fswb = new FsWorkbook();
            const fsws = fswb.InitWorksheet("My Awesome Worksheet");
            fsws.Row(1).Item(2).SetValueAs("My Column 1");
            fsws.Row(2).Item(2).SetValueAs(2);
            fsws.Row(3).Item(2).SetValueAs(20);
            fsws.Row(1).Item(3).SetValueAs("My Column 2");
            fsws.Row(2).Item(3).SetValueAs("row2");
            fsws.Row(3).Item(3).SetValueAs("row20");
            fsws.Row(1).Item(4).SetValueAs("My Column 3");
            fsws.Row(2).Item(4).SetValueAs(true);
            fsws.Row(3).Item(4).SetValueAs(false);
            const table = new FsTable("MyNewTable", FsRangeAddress_$ctor_Z721C83C5("B1:D3"));
            fsws.AddTable(table);
            fsws.RescanRows()
            const jswb = toJsWorkbook(fswb)
            await jswb.xlsx.writeFile(path);
            let readjswb = new Excel.Workbook();
            await readjswb.xlsx.readFile(path)
            let readfswb = toFsWorkbook(readjswb)
            equal(readfswb.GetWorksheets().length, fswb.GetWorksheets().length)
            equal(readfswb.GetWorksheets()[0].Name, "My Awesome Worksheet")
            equal(readfswb.GetWorksheets()[0].Name, fswb.GetWorksheets()[0].Name)
            equal(readfswb.GetWorksheets()[0].Tables.length, fswb.GetWorksheets()[0].Tables.length)
            equal(readfswb.GetWorksheets()[0].Tables[0].Name, fswb.GetWorksheets()[0].Tables[0].Name)
            equal(readfswb.GetWorksheets()[0].GetCellAt(1,2).Value, "My Column 1")
            equal(readfswb.GetWorksheets()[0].GetCellAt(2,2).ValueAsFloat(), 2)
            equal(readfswb.GetWorksheets()[0].GetCellAt(2,3).Value, "row2")
            equal(readfswb.GetWorksheets()[0].GetCellAt(2,4).ValueAsBool(), true)
            const readwritejswb = toJsWorkbook(readfswb)
            await readwritejswb.xlsx.writeFile(path2);
        })
        // it('combined function', async () => {
        //     const path = "tests/JS/TestFiles/ReadTable.xlsx";
        //     console.log("start")
        //     const wb = await Xlsx.fromXlsxFile(path)
        //     console.log(wb)
        // });
    });
});