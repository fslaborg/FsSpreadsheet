module FsTable

open Expecto
open FsSpreadsheet


let dummyFsCells = 
    [   // rows
        [   // single cells in row
            FsCell.createWithDataType DataType.String 2 2 "Name"
            FsCell.createWithDataType DataType.String 2 3 "Age"
            FsCell.createWithDataType DataType.String 2 4 "Location"
        ]
        [
            FsCell.createWithDataType DataType.String 3 2 "John Doe"
            FsCell.createWithDataType DataType.Number 3 3 "69"
            FsCell.createWithDataType DataType.String 3 4 "Springfield"
        ]
        [
            FsCell.createWithDataType DataType.String 4 2 "Jane Doe"
            FsCell.createWithDataType DataType.Number 4 3 "23"
            FsCell.createWithDataType DataType.String 4 4 "Springfield"
        ]
        [
            FsCell.createWithDataType DataType.String 5 2 "Jack Doe"
            FsCell.createWithDataType DataType.Number 5 3 "4"
            FsCell.createWithDataType DataType.String 5 4 "Newville"
        ]
    ]
    |> Seq.concat
let dummyFsRangeColumns =
    dummyFsCells
    |> Seq.groupBy (fun t -> t.ColumnNumber)
    |> Seq.map (
        snd 
        >> fun fscs -> 
            (
                fscs 
                |> Seq.minBy (fun fsc -> fsc.RowNumber) 
                |> fun fsc -> FsAddress(fsc.RowNumber, fsc.ColumnNumber)
                ,
                fscs 
                |> Seq.maxBy (fun fsc -> fsc.RowNumber) 
                |> fun fsc -> FsAddress(fsc.RowNumber, fsc.ColumnNumber)
            )
            |> FsRangeAddress
        >> FsRangeColumn
    )
let dummyFsCellsCollection = FsCellsCollection()
dummyFsCellsCollection.Add dummyFsCells |> ignore
let dummyFsCellsCollectionFirstAddress = dummyFsCellsCollection.GetFirstAddress()
let dummyFsCellsCollectionLastAddress = dummyFsCellsCollection.GetLastAddress()
let dummyFsTable = FsTable("dummyFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))
let dummyFsTableFields =
    dummyFsCells
    |> Seq.map (
        fun fsc -> 
            FsTableField(
                fsc.Value, 
                fsc.ColumnNumber, 
                dummyFsRangeColumns |> Seq.find (fun t -> t.Index = fsc.ColumnNumber), 
                obj, 
                obj
            )
    )


[<Tests>]
let fsTableTests =
    testList "FsTable" [
        testList "Fields" [
            testCase "Correctly retrieved" <| fun _ ->
                let actualFields = dummyFsTable.Fields dummyFsCellsCollection
                Expect.containsAll actualFields dummyFsTableFields "FsTableFields do not match"
        ]
    ]