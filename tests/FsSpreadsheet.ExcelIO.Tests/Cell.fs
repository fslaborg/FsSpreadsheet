module Cell

open Expecto
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml


let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "data", "testUnit.xlsx")
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.ExcelIO
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let sstpFox = wbpFox.SharedStringTablePart
let sstFox = sstpFox.SharedStringTable
let sstFoxInnerText = sstFox.InnerText
let wsp1Fox = (wbpFox.WorksheetParts |> Array.ofSeq)[0]
let cbsi1Fox = wsp1Fox.Worksheet.Descendants<Spreadsheet.Cell>() |> Array.ofSeq
let nullCell = Cell.create Spreadsheet.CellValues.Error "A1" (Cell.CellValue.create "")
nullCell.CellValue.Text <- null


[<Tests>]
let cellTests =
    testList "Cell" [
        testList "includeSharedStringValue" [
            let cissv1_0 = Cell.includeSharedStringValue sstFox cbsi1Fox[0]
            testCase "element 0 with included SharedStringValue is equal in CellValueText to element 0 from OpenXML" <| fun _ ->
                Expect.equal cissv1_0.CellValue.Text cbsi1Fox[0].CellValue.Text "Differs"
            let cissvNull = Cell.includeSharedStringValue sstFox nullCell
            testCase "nullCell with included SharedStringValue has no CellValueText" <| fun _ ->
                Expect.notEqual cissv1_0.CellValue.Text cissvNull.CellValue.Text "Does not differ"
        ]
    ]