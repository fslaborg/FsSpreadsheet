module FsWorkbook

open Expecto
open FsSpreadsheet


let dummyWorkbook = new FsWorkbook()
let dummyWorksheet = FsWorksheet("dummyWorksheet")
dummyWorkbook.AddWorksheet dummyWorksheet |> ignore


[<Tests>]
let fsWorkbookTests =
    testList "FsWorkbook" [
        testList "TryGetWorksheetByName" [
            let testWorksheet = dummyWorkbook.TryGetWorksheetByName "dummyWorksheet"
            testCase "is Some" <| fun _ ->
                Expect.isSome testWorksheet "is None"
            // TO DO: add more cases
        ]
        testList "GetWorksheetByName" [
            testCase "does not throw exception when present" <| fun _ ->
                Expect.isOk (dummyWorkbook.GetWorksheetByName "dummyWorksheet" |> Result.Ok) "did throw exception"
            testCase "does throw exception when not present" <| fun _ ->
                Expect.throws (fun () -> dummyWorkbook.GetWorksheetByName "bla" |> ignore) "did not throw exception"
        ]
    ]