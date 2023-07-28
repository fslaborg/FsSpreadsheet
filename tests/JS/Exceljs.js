import { equal } from 'assert';
import { Xlsx } from './FsSpreadsheet.Exceljs/Xlsx.fs.js';
import { Excel } from './FsSpreadsheet.Exceljs/fable_modules/Fable.Exceljs.1.3.6/ExcelJs.fs.js';
import { toJsWorkbook, toFsWorkbook } from "./FsSpreadsheet.Exceljs/Workbook.fs.js";

describe('FsSpreadsheet.Exceljs', function () {
    describe('read', function () {
        it('table', async () => {
            //const wb = await Xlsx.fromXlsxFile(path)
            const path = "C:/Users/Kevin/source/repos/FsSpreadsheet/tests/JS/TestFiles/ReadTable.xlsx";
            const wb = new Excel.Workbook();
            await wb.xlsx.readFile(path)
            const fswb = toFsWorkbook(wb)
            console.log(fswb)
            let worksheets = fswb.GetWorksheets()
            equal(worksheets, 1)
        });
    });
});