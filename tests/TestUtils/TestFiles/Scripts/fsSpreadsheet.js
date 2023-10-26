import { Xlsx } from "./fable/Xlsx.js"

export const inputPath = "../TestWorkbook_Excel.xlsx";

export const outputPath = "../TestWorkbook_FsSpreadsheet.js.xlsx";

async function run() {
    let wb = await Xlsx.fromXlsxFile(inputPath)
    // console.log(wb)
    Xlsx.toFile(outputPath, wb)
}

run();

