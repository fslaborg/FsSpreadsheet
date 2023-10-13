import { PromiseBuilder__Delay_62FBFDE1, PromiseBuilder__Run_212F1D4B } from "./fable_modules/Fable.Promise.3.2.0/Promise.fs.js";
import { Excel } from "./fable_modules/Fable.Exceljs.1.6.0/ExcelJs.fs.js";
import { promise } from "./fable_modules/Fable.Promise.3.2.0/PromiseImpl.fs.js";

export const inputPath = "../TestWorkbook_Excel.xlsx";

export const outputPath = "../TestWorkbook_FableExceljs.xlsx";

export function run() {
    return PromiseBuilder__Run_212F1D4B(promise, PromiseBuilder__Delay_62FBFDE1(promise, () => {
        const wb = new Excel.Workbook();
        return wb.xlsx.readFile(inputPath).then(() => (wb.xlsx.writeFile(outputPath)));
    }));
}

run();

