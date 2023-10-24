import { equal } from 'assert';
import { Xlsx } from './FsSpreadsheet.Exceljs/Xlsx.js';
import { FsWorkbook } from "./FsSpreadsheet.Exceljs/FsSpreadsheet/FsWorkbook.js";
import { FsRangeAddress_$ctor_Z721C83C5, FsRangeAddress__get_Range } from "./FsSpreadsheet.Exceljs/FsSpreadsheet/Ranges/FsRangeAddress.js";
import { FsTable } from "./FsSpreadsheet.Exceljs/FsSpreadsheet/Tables/FsTable.js";
import { writeFromFsWorkbook, readToFsWorkbook } from "./FsSpreadsheet.Exceljs/Workbook.js";

describe('FsSpreadsheet.Exceljs', function () {
    describe('read', function () {
        it('values', async () => {
            const path = "tests/JS/TestFiles/ReadTable.xlsx"; // path always from package.json 
            const fswb = await Xlsx.fromXlsxFile(path)
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
            const path = "tests/JS/TestFiles/ReadTable.xlsx";
            const fswb = await Xlsx.fromXlsxFile(path)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
            let ws = worksheets[0]
            equal(ws.Name, "Tabelle1")
            let tables = ws.Tables
            equal(tables.length, 1)
            let table = ws.Tables[0]
            equal(table.Name, "Table1")
        });
        it('isa.investigation.xlsx', async () => {
            const path = "tests/JS/TestFiles/isa.investigation.xlsx";
            const fswb = await Xlsx.fromXlsxFile(path)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
            let ws = worksheets[0]
            equal(ws.Name, "isa_investigation")
        });
        it('isa_assay_keineTables', async () => {
            const path = "tests/JS/TestFiles/isa_assay_keineTables.xlsx";
            const fswb = await Xlsx.fromXlsxFile(path)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
        });
        // issue #58
        // TypeError: Cannot read properties of undefined (reading 'company')
        // it('ClosedXml.Table', async () => {
        //    const path = "tests/JS/TestFiles/ClosedXml.Table.xlsx";
        //    const fswb = await Xlsx.fromXlsxFile(path)
        //    let worksheets = fswb.GetWorksheets()
        //    equal(worksheets.length, 1)
        // });
        it('fsspreadsheet.minimalTable', async () => {
            const path = "tests/JS/TestFiles/fsspreadsheet.minimalTable.xlsx";
            const fswb = await Xlsx.fromXlsxFile(path)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
        });
        it('isa.study.xlsx', async () => {
            const path = "tests/JS/TestFiles/isa.study.xlsx";
            const fswb = await Xlsx.fromXlsxFile(path)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 5)
        });
        it('readOldClosedXml', async () => {
            const path = "tests/JS/TestFiles/readOldClosedXml.xlsx";
            const fswb = await Xlsx.fromXlsxFile(path)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets.length, 1)
        });
    })
    describe('write', function () {
        it('passes', async () => {
            const path = "tests/JS/TestFiles/WRITE_Table.xlsx"
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
            await Xlsx.toFile(path, fswb)
            const readfswb = await Xlsx.fromXlsxFile(path)
            equal(readfswb.GetWorksheets().length, fswb.GetWorksheets().length)
            equal(readfswb.GetWorksheets()[0].Name, "My Awesome Worksheet")
            equal(readfswb.GetWorksheets()[0].Name, fswb.GetWorksheets()[0].Name)
            equal(readfswb.GetWorksheets()[0].Tables.length, fswb.GetWorksheets()[0].Tables.length)
            equal(readfswb.GetWorksheets()[0].Tables[0].Name, fswb.GetWorksheets()[0].Tables[0].Name)
            equal(readfswb.GetWorksheets()[0].GetCellAt(1,2).Value, "My Column 1")
            equal(readfswb.GetWorksheets()[0].GetCellAt(2,2).ValueAsFloat(), 2)
            equal(readfswb.GetWorksheets()[0].GetCellAt(2,3).Value, "row2")
            equal(readfswb.GetWorksheets()[0].GetCellAt(2,4).ValueAsBool(), true)
        })
    });
    describe('read-write', function (){
        it('from excel', async () => {
            const inputPath = "tests/JS/TestFiles/TestAssayExcel.xlsx"
            const fswb = await Xlsx.fromXlsxFile(inputPath)
            equal(fswb.GetWorksheets().length, 5, "test correct read")
            let ws1 = fswb.GetWorksheets()[0]
            equal(ws1.name, "Cell Lysis", "ws1.name")
            let table = ws1.Tables[0]
            equal(table.Name, "annotationTableStupidQuail41", "table.Name")
            equal(table.ShowHeaderRow, true, "table.ShowHeaderRow") // issue #69
            const outoutPath = "tests/JS/TestFiles/WRITE_TestAssayExcel.xlsx"
            await Xlsx.toFile(outoutPath, fswb)
            const fswb2 = await Xlsx.fromXlsxFile(outoutPath)
            equal(fswb2.GetWorksheets().length, 5) // test correct read
        })
    })
});