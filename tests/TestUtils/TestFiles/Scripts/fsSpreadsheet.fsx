#r "nuget: FsSpreadsheet.ExcelIO"

open FsSpreadsheet
open FsSpreadsheet.ExcelIO

let inputPath = @"../TestWorkbook_Excel.xlsx"

let outputPath = @"../TestWorkbook_FsSpreadsheet.xlsx"

let wb = FsWorkbook.fromXlsxFile (inputPath)

wb.ToFile(outputPath)
