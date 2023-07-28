module FsWorkSheet

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif
open FsSpreadsheet

let dummyCellsColl = FsCellsCollection()
let dummyTable1 = FsTable("dummyTable1", FsRangeAddress("A1:B2"))
let dummyTable2 = FsTable("dummyTable2", FsRangeAddress("D1:F3"))
let dummySheet1 = FsWorksheet("dummySheet1", ResizeArray(), ResizeArray(), dummyCellsColl)
let dummySheet2 = FsWorksheet("dummySheet2", ResizeArray(), ResizeArray([dummyTable1; dummyTable2]), dummyCellsColl)

let main =
    testSequenced <| testList "FsWorksheet" [
        testList "FsCell data" [
            // TO DO: Ask TM: useful? or was that a mistake? (since the same test is seen in FsCell.fs)
            testList "Data | DataType | Adress" [
                let fscellA1_string  = FsCell.create 1 1 "A1"
                let fscellB1_num     = FsCell.create 1 2 1
                let fscellA2_bool    = FsCell.create 1 2 true
                //let worksheet = FsWorksheet.
                testCase "DataType string" <| fun _ ->
                    Expect.equal fscellA1_string.DataType DataType.String "is not the expected DataType.String"
            ]
        ]
        testList "FsTable methods" [
            testList "tryGetTableByName" [
                let testTableOption = FsWorksheet.tryGetTableByName "dummyTable1" dummySheet2
                testCase "is Some" <| fun _ ->
                    Expect.isSome testTableOption "is None"
                // TO DO: add more testCases
            ]
            testList "getTableByName" [
                let testTable = FsWorksheet.getTableByName "dummyTable1" dummySheet2
                testCase "is equal to dummyTable1" <| fun _ ->
                    Expect.equal testTable dummyTable1 "is not equal"
                // TO DO: add more testCases
            ]
            testList "AddTable" [
                testCase "dummyTable1 is present" <| fun _ ->
                    let ws = new FsWorksheet("testSheet", ResizeArray(), ResizeArray(), FsCellsCollection())
                    ws.AddTable dummyTable1 |> ignore
                    Expect.myContains ws.Tables dummyTable1 "does not contain dummyTable1"
                testCase "dummyTable1 is not present twice" <| fun _ ->
                    let ws = new FsWorksheet("testSheet", ResizeArray(), ResizeArray(), FsCellsCollection())
                    ws.AddTable dummyTable1 |> ignore
                    ws.AddTable dummyTable1 |> ignore
                    Expect.myHasCountOf ws.Tables 1 (fun t -> t.Name = dummyTable1.Name) "has dummyTable1 twice (or more)"
                // DO DO: add more testCases
            ]
            testList "AddTables" [
                testCase "dummyTable1 is present" <| fun _ ->
                    let testSheet = FsWorksheet("testSheet", ResizeArray(), ResizeArray(), FsCellsCollection())
                    testSheet.AddTables [dummyTable1] |> ignore
                    Expect.myContains testSheet.Tables dummyTable1 "does not contain dummyTable1"
                testCase "dummyTable1 & dummyTable2 are present" <| fun _ ->
                    let testSheet = FsWorksheet("testSheet", ResizeArray(), ResizeArray(), FsCellsCollection())
                    let dummyTablesList = [dummyTable1; dummyTable2]
                    testSheet.AddTables dummyTablesList |> ignore
                    Expect.containsAll testSheet.Tables dummyTablesList "does not contain dummyTable1 and/or dummyTable2"
                testCase "dummyTable1 & dummyTable2 are not present twice" <| fun _ ->
                    let testSheet = FsWorksheet("testSheet", ResizeArray(), ResizeArray(), FsCellsCollection())
                    let dummyTablesList = [dummyTable1; dummyTable2]
                    testSheet.AddTables dummyTablesList |> ignore
                    testSheet.AddTables dummyTablesList |> ignore
                    Expect.containsAll testSheet.Tables dummyTablesList "has dummyTable1 and/or twice (or more)"
            ]
        ]
    ]