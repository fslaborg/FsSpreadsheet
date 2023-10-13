#r "nuget: ClosedXML, 0.102.1"

open ClosedXML
open ClosedXML.Excel

let inputPath = @"../TestWorkbook_Excel.xlsx"

let outputPath = @"../TestWorkbook_ClosedXML.xlsx"

let wb = new XLWorkbook(inputPath)

wb.SaveAs(outputPath)
