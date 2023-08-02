import { FsWorkbook } from "./FsSpreadsheet/FsWorkbook.js";
import { addFsWorksheet, addJsWorksheet } from "./Worksheet.js";
import { Excel } from "./fable_modules/Fable.Exceljs.1.5.0/ExcelJs.fs.js";
import { disposeSafe, getEnumerator } from "./fable_modules/fable-library.4.1.3/Util.js";

export function toFsWorkbook(jswb) {
    const fswb = new FsWorkbook();
    const arr = jswb.worksheets;
    for (let idx = 0; idx <= (arr.length - 1); idx++) {
        addJsWorksheet(fswb, arr[idx]);
    }
    return fswb;
}

export function toJsWorkbook(fswb) {
    const jswb = new Excel.Workbook();
    let enumerator = getEnumerator(fswb.GetWorksheets());
    try {
        while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
            addFsWorksheet(jswb, enumerator["System.Collections.Generic.IEnumerator`1.get_Current"]());
        }
    }
    finally {
        disposeSafe(enumerator);
    }
    return jswb;
}

