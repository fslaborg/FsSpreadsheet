#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Debug\net6.0\DocumentFormat.OpenXml.dll"
#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Debug\net6.0\FsSpreadsheet.dll"
#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Debug\net6.0\FsSpreadsheet.ExcelIO.dll"
#r @"..\..\..\FsSpreadsheet.ExcelIO.Tests\bin\Debug\net6.0\System.IO.Packaging.dll"


open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml

let inputPath = @"../TestWorkbook_Excel.xlsx"

let outputPath = @"../TestWorkbook_FsSpreadsheet.net.xlsx"

let wb = FsWorkbook.fromXlsxFile (inputPath)
// wb.GetWorksheets().[0].GetCellAt(5,1) |> fun x -> (x.Value, x.DataType) |> printfn "%A"

// let r = wb.GetWorksheets().[0].GetCellAt(5,1).Value |> string

// Cell(DataType = EnumValue(CellValues.Number), CellValue = CellValue(r)).InnerText

// for i in r do printfn "%A" i

wb.ToFile(outputPath)
