module Workbook

open Expecto
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml

let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "../data", "testUnit.xlsx")
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.ExcelIO
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let wbFox = wbpFox.Workbook


[<Tests>]
let workbookTests =
    testList "Workbook" [
        testList "get" [
            testCase "is equal to wbFox" <| fun _ ->
                let wb = Workbook.get wbpFox
                Expect.equal wb wbFox "Differs"
        ]
    ]