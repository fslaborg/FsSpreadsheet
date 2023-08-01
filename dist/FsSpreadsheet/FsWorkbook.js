import { collect, tryFind as tryFind_1, removeInPlace, map } from "../fable_modules/fable-library.4.1.3/Array.js";
import { FsWorksheet } from "./FsWorksheet.js";
import { tryFind, tryItem, iterate, exists } from "../fable_modules/fable-library.4.1.3/Seq.js";
import { printf, toFail } from "../fable_modules/fable-library.4.1.3/String.js";
import { defaultArg, value as value_1 } from "../fable_modules/fable-library.4.1.3/Option.js";
import { safeHash, equals, defaultOf } from "../fable_modules/fable-library.4.1.3/Util.js";
import { class_type } from "../fable_modules/fable-library.4.1.3/Reflection.js";

export class FsWorkbook {
    constructor() {
        this._worksheets = [];
    }
    Copy() {
        const self = this;
        const shts = map((s) => s.Copy(), self.GetWorksheets().slice());
        const wb = new FsWorkbook();
        wb.AddWorksheets(shts);
        return wb;
    }
    static copy(workbook) {
        return workbook.Copy();
    }
    InitWorksheet(name) {
        const self = this;
        const sheet = new FsWorksheet(name);
        void (self._worksheets.push(sheet));
        return sheet;
    }
    static initWorksheet(name, workbook) {
        return workbook.InitWorksheet(name);
    }
    AddWorksheet(sheet) {
        const self = this;
        if (exists((ws) => (ws.Name === sheet.Name), self._worksheets)) {
            const arg = sheet.Name;
            toFail(printf("Could not add worksheet with name \"%s\" to workbook as it already contains a worksheet with the same name"))(arg);
        }
        else {
            void (self._worksheets.push(sheet));
        }
    }
    static addWorksheet(sheet, workbook) {
        workbook.AddWorksheet(sheet);
        return workbook;
    }
    AddWorksheets(sheets) {
        const self = this;
        iterate((arg) => {
            self.AddWorksheet(arg);
        }, sheets);
    }
    static addWorksheets(sheets, workbook) {
        workbook.AddWorksheets(sheets);
        return workbook;
    }
    GetWorksheets() {
        const self = this;
        return self._worksheets;
    }
    static getWorksheets(workbook) {
        return workbook.GetWorksheets();
    }
    TryGetWorksheetAt(index) {
        const self = this;
        return tryItem(index - 1, self._worksheets);
    }
    static tryGetWorksheetAt(index, workbook) {
        return workbook.TryGetWorksheetAt(index);
    }
    GetWorksheetAt(index) {
        const self = this;
        const matchValue = self.TryGetWorksheetAt(index);
        if (matchValue == null) {
            throw new Error(`FsWorksheet at position ${index} is not present in the FsWorkbook.`);
        }
        else {
            return matchValue;
        }
    }
    static getWorksheetAt(index, workbook) {
        return workbook.GetWorksheetAt(index);
    }
    TryGetWorksheetByName(sheetName) {
        const self = this;
        return tryFind((w) => (w.Name === sheetName), self._worksheets);
    }
    static tryGetWorksheetByName(sheetName, workbook) {
        return workbook.TryGetWorksheetByName(sheetName);
    }
    GetWorksheetByName(sheetName) {
        const self = this;
        try {
            return value_1(self.TryGetWorksheetByName(sheetName));
        }
        catch (matchValue) {
            throw new Error(`FsWorksheet with name ${sheetName} is not present in the FsWorkbook.`);
        }
    }
    static getWorksheetByName(sheetName, workbook) {
        return workbook.GetWorksheetByName(sheetName);
    }
    RemoveWorksheet(name) {
        const self = this;
        removeInPlace((() => {
            try {
                return defaultArg(tryFind_1((ws) => (ws.Name === name), self._worksheets), defaultOf());
            }
            catch (matchValue) {
                throw new Error(`FsWorksheet with name ${name} was not found in FsWorkbook.`);
            }
        })(), self._worksheets, {
            Equals: equals,
            GetHashCode: safeHash,
        });
    }
    static removeWorksheet(name, workbook) {
        workbook.RemoveWorksheet(name);
        return workbook;
    }
    GetTables() {
        const self = this;
        return collect((s) => Array.from(s.Tables), self.GetWorksheets().slice());
    }
    static getTables(workbook) {
        return workbook.GetTables();
    }
    Dispose() {
    }
}

export function FsWorkbook_$reflection() {
    return class_type("FsSpreadsheet.FsWorkbook", void 0, FsWorkbook);
}

export function FsWorkbook_$ctor() {
    return new FsWorkbook();
}

