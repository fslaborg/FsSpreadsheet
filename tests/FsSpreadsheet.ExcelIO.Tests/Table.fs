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
    
    ]
    

[<Tests>]
let main =
    testList "FsTable" [
        transformTable
    ]
