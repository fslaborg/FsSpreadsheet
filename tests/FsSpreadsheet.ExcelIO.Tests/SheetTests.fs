module Sheet

open Expecto
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml

let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "data", "testUnit.xlsx")
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.ExcelIO
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let wbFox = wbpFox.Workbook
let shtsFox = wbFox.Sheets
let shtssFox = shtsFox.Descendants<Spreadsheet.Sheet>() |> Array.ofSeq      // array is needed since seqs cannot be compared


[<Tests>]
let sheetsTests =
    testList "Sheets" [
        testList "get" [
            testCase "is equal to shtsFox" <| fun _ ->
                let shts = Sheet.Sheets.get wbFox
                Expect.equal shts shtsFox "Differs"
        ]
        testList "getSheets" [
            testCase "is equal to shtssFox" <| fun _ ->
                let shtss = Sheet.Sheets.getSheets shtsFox |> Array.ofSeq   // array is needed since seqs cannot be compared
                Expect.equal shtss shtssFox "Differs"
        ]
    ]