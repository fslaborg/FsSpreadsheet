module FsTable

open Expecto
open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml
open TestingUtils


let transformTable =
    testList "transformTable" [
        testCase "handleNullFields" (fun () ->
            let table = FsSpreadsheet.ExcelIO.Table.create "TestTable" (StringValue ("A1:D4")) []
            Expect.isTrue (table.TotalsRowShown = null) "Check that field of interest is None"
            FsTable.fromXlsxTable table |> ignore
            )
        testCase "handleNoHeaders" (fun () ->
            let doc = Spreadsheet.fromFile TestObjects.headerLessTablePath false
            let wsp = doc.WorkbookPart.WorksheetParts |> Seq.head
            let t = wsp |> Worksheet.WorksheetPart.getTables |> Seq.head
            let parsed = FsTable.fromXlsxTable t
            Expect.equal (parsed.RangeAddress.ToString()) "B8:B8" "Range address should be B8:B8"
            Expect.isFalse parsed.ShowHeaderRow "ShowHeaderRow should be false"
        )
    ]
    

[<Tests>]
let main =
    testList "FsTable" [
        transformTable
    ]
