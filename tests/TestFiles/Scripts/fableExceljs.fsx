#r "nuget: Fable.Exceljs, 1.6.0"
#r "nuget: Fable.Promise, 3.2.0"

open Fable.ExcelJs

let inputPath = @"../TestWorkbook_Excel.xlsx"

let outputPath = @"../TestWorkbook_FableExceljs.xlsx"


let run() =
    promise {
        let wb = ExcelJs.Excel.Workbook()
        // Read
        do! wb.xlsx.readFile(inputPath)
        // Write
        return! wb.xlsx.writeFile(outputPath)
    }

run()
