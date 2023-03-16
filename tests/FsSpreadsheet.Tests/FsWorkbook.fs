module FsWorkbook

open Expecto
open FsSpreadsheet


let dummyWorkbook = new FsWorkbook()
let dummyWorksheet1 = FsWorksheet("dummyWorksheet1")
let dummyWorksheet2 = FsWorksheet("dummyWorksheet2")
let dummyTables = [
    FsTable("dummyTable1", FsRangeAddress("A1:B2"))
    FsTable("dummyTable2", FsRangeAddress("C3:F5"))
]
dummyWorkbook.AddWorksheet dummyWorksheet1 |> ignore
dummyWorkbook.AddWorksheet dummyWorksheet2 |> ignore
dummyWorksheet1.AddTable dummyTables[0] |> ignore
dummyWorksheet2.AddTable dummyTables[1] |> ignore


[<Tests>]
let fsWorkbookTests =
    testList "FsWorkbook" [
        testList "TryGetWorksheetByName" [
            let testWorksheet = dummyWorkbook.TryGetWorksheetByName "dummyWorksheet1"
            testCase "is Some" <| fun _ ->
                Expect.isSome testWorksheet "is None"
            // TO DO: add more cases
        ]
        testList "GetWorksheetByName" [
            testCase "does not throw exception when present" <| fun _ ->
                Expect.isOk (dummyWorkbook.GetWorksheetByName "dummyWorksheet1" |> Result.Ok) "did throw exception"
            testCase "does throw exception when not present" <| fun _ ->
                Expect.throws (fun () -> dummyWorkbook.GetWorksheetByName "bla" |> ignore) "did not throw exception"
        ]
        testList "GetTables" [
            testCase "got correct FsTables" <| fun _ ->
                let tablesGotten = dummyWorkbook.GetTables()
                Expect.equal tablesGotten dummyTables "Did not get correct FsTables"
        ]
    ]