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
            FsCell.createWithDataType DataType.String 3 3 "Springfield"
        ]
        [
            FsCell.createWithDataType DataType.String 4 2 "Jane Doe"
            FsCell.createWithDataType DataType.Number 4 3 "23"
            FsCell.createWithDataType DataType.String 4 3 "Springfield"
        ]
        [
            FsCell.createWithDataType DataType.String 5 2 "Jack Doe"
            FsCell.createWithDataType DataType.Number 5 3 "4"
            FsCell.createWithDataType DataType.String 5 3 "Newville"
        ]
    ]
let dummyFsCellsCollection = FsCellsCollection()
dummyFsCellsCollection.Add(List.concat dummyFsCells) |> ignore
let dummyFsCellsCollectionFirstAddress = dummyFsCellsCollection.GetFirstAddress()
let dummyFsCellsCollectionLastAddress = dummyFsCellsCollection.GetLastAddress()
let dummyFsTable = FsTable("dummyFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))


[<Tests>]
let fsTableTests =
    testList "FsTable" [
        testList "" [
            
        ]
    ]