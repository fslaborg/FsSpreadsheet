module FsWorkSheet
open Expecto
open FsSpreadsheet


let dummyCellsColl = FsCellsCollection()
let dummyTable1 = FsTable("dummyTable1", FsRangeAddress("A1:B2"))
let dummyTable2 = FsTable("dummyTable2", FsRangeAddress("D1:F3"))
let dummySheet1 = FsWorksheet("dummySheet1", [], [], dummyCellsColl)
let dummySheet2 = FsWorksheet("dummySheet2", [], [dummyTable1; dummyTable2], dummyCellsColl)



[<Tests>]
let fsWorksheetTest =
    testList "FsWorksheet" [
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
                let testSheet = FsWorksheet("testSheet", [], [], FsCellsCollection())
                testSheet.AddTable dummyTable1 |> ignore
                testCase "dummyTable1 is present" <| fun _ ->
                    Expect.contains testSheet.Tables dummyTable1 "does not contain dummyTable1"
                testCase "dummyTable1 is not present twice" <| fun _ ->
                    testSheet.AddTable dummyTable1 |> ignore
                    Expect.hasCountOf testSheet.Tables 1u (fun t -> t.Name = dummyTable1.Name) "has dummyTable1 twice (or more)"
                // DO DO: add more testCases
            ]
        ]
    ]