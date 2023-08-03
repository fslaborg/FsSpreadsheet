import { PromiseBuilder__Delay_62FBFDE1, PromiseBuilder__Run_212F1D4B } from "./fable_modules/Fable.Promise.3.2.0/Promise.fs.js";
import { Excel } from "./fable_modules/Fable.Exceljs.1.6.0/ExcelJs.fs.js";
import { promise } from "./fable_modules/Fable.Promise.3.2.0/PromiseImpl.fs.js";
import { toJsWorkbook, toFsWorkbook } from "./Workbook.js";
import { class_type } from "./fable_modules/fable-library.4.1.3/Reflection.js";

export class Xlsx {
    constructor() {
    }
    static fromXlsxFile(path) {
        return PromiseBuilder__Run_212F1D4B(promise, PromiseBuilder__Delay_62FBFDE1(promise, () => {
            const wb = new Excel.Workbook();
            return wb.xlsx.readFile(path).then(() => {
                const fswb = toFsWorkbook(wb);
                return Promise.resolve(fswb);
            });
        }));
    }
    static fromXlsxStream(stream) {
        return PromiseBuilder__Run_212F1D4B(promise, PromiseBuilder__Delay_62FBFDE1(promise, () => {
            const wb = new Excel.Workbook();
            return wb.xlsx.read(stream).then(() => (Promise.resolve(toFsWorkbook(wb))));
        }));
    }
    static fromBytes(bytes) {
        return PromiseBuilder__Run_212F1D4B(promise, PromiseBuilder__Delay_62FBFDE1(promise, () => {
            const wb = new Excel.Workbook();
            const uint8 = new Uint8Array(bytes);
            return wb.xlsx.load(uint8.buffer).then(() => (Promise.resolve(toFsWorkbook(wb))));
        }));
    }
    static toFile(path, wb) {
        const jswb = toJsWorkbook(wb);
        return jswb.xlsx.writeFile(path);
    }
    static toStream(stream, wb) {
        const jswb = toJsWorkbook(wb);
        return jswb.xlsx.write(stream);
    }
    static toBytes(wb) {
        return PromiseBuilder__Run_212F1D4B(promise, PromiseBuilder__Delay_62FBFDE1(promise, () => {
            const jswb = toJsWorkbook(wb);
            const buffer = jswb.xlsx.writeBuffer();
            return Promise.resolve(buffer);
        }));
    }
}

export function Xlsx_$reflection() {
    return class_type("FsSpreadsheet.Exceljs.Xlsx", void 0, Xlsx);
}

