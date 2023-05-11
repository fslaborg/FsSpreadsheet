module FsWorkbook

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif
open FsSpreadsheet


let dummyWorkbook = new FsWorkbook()
let dummyWorksheet1 = FsWorksheet("dummyWorksheet1")
let dummyWorksheet2 = FsWorksheet("dummyWorksheet2")
let dummyWorksheetList = [dummyWorksheet1; dummyWorksheet2]
let dummyTables = [
    FsTable("dummyTable1", FsRangeAddress("A1:B2"))
    FsTable("dummyTable2", FsRangeAddress("C3:F5"))
]
dummyWorkbook.AddWorksheet dummyWorksheet1 |> ignore
dummyWorkbook.AddWorksheet dummyWorksheet2 |> ignore
dummyWorksheet1.AddTable dummyTables[0] |> ignore
dummyWorksheet2.AddTable dummyTables[1] |> ignore


let main =
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
            testCase "gets correct FsTables" <| fun _ ->
                let tablesGotten = dummyWorkbook.GetTables()
                Expect.equal tablesGotten dummyTables "Did not get correct FsTables"
        ]
        testList "AddWorksheets" [
            testCase "adds all FsWorksheets correctly" <| fun _ ->
                let testWorkbook = new FsWorkbook()
                testWorkbook.AddWorksheets [dummyWorksheet1; dummyWorksheet2] |> ignore
                let testWorkbookWorksheetNames = testWorkbook.GetWorksheets() |> List.map (fun ws -> ws.Name)
                let dummyWorksheetNames = dummyWorksheetList |> List.map (fun ws -> ws.Name)
                Expect.containsAll testWorkbookWorksheetNames dummyWorksheetNames "Does not contain all FsWorksheets"
        ]
    ]