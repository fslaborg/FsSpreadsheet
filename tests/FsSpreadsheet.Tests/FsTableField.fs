module FsTableField

open FsSpreadsheet
open Expecto


let dummyFsRangeAddress = FsRangeAddress("A1:A3")
let dummyFsRangeColumn = FsRangeColumn(dummyFsRangeAddress)
let dummyFsTableField = FsTableField("dummyFsTableField", 0, dummyFsRangeColumn, obj, obj)
let dummyFsCells = [
    FsCell.createWithDataType DataType.String 1 1 "I am the Header!"
    FsCell.createWithDataType DataType.String 2 1 "first data cell"
    FsCell.createWithDataType DataType.String 3 1 "second data cell"
]
let dummyFsCellsCollection = FsCellsCollection()
dummyFsCellsCollection.Add dummyFsCells |> ignore


[<Tests>]
let fsTableField =
    testList "FsTableField" [
        testList "" [
            testCase "" <| fun _ ->
                Expect.equal  "not equal"
        ]
    ]