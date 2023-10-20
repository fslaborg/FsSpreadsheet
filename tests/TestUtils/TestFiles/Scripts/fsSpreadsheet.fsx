#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Release\net6.0\DocumentFormat.OpenXml.dll"
#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Release\net6.0\FsSpreadsheet.dll"
#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Release\net6.0\FsSpreadsheet.ExcelIO.dll"
#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Release\net6.0\System.IO.Packaging.dll"


open FsSpreadsheet
open FsSpreadsheet.ExcelIO

let inputPath = @"../TestWorkbook_Excel.xlsx"

let outputPath = @"../TestWorkbook_FsSpreadsheet.net.xlsx"

let wb = FsWorkbook.fromXlsxFile (inputPath)

wb.ToFile(outputPath)
