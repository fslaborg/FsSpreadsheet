module Workbook.Tests

open TestingUtils
open FsSpreadsheet.Net
open DocumentFormat.OpenXml

let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "../data", "testUnit.xlsx")
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.Net
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let wbFox = wbpFox.Workbook


let workbookTests =
    testList "Workbook" [
        testList "get" [
            testCase "is equal to wbFox" <| fun _ ->
                let wb = Workbook.get wbpFox
                Expect.equal wb wbFox "Differs"
        ]
    ]

let main = workbookTests