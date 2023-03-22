module FsCellsCollection

open Expecto
open FsSpreadsheet


let dummyFsCells =
    [
        [
            FsCell.createWithDataType DataType.String 1 1 "A1"
            FsCell.createWithDataType DataType.String 1 2 "B1"
            FsCell.createWithDataType DataType.String 1 3 "C1"
        ]
        [
            FsCell.createWithDataType DataType.String 2 1 "A2"
            FsCell.createWithDataType DataType.String 2 2 "B2"
            FsCell.createWithDataType DataType.String 2 3 "C2"
        ]
        [
            FsCell.createWithDataType DataType.String 3 1 "A3"
            FsCell.createWithDataType DataType.String 3 2 "B3"
            FsCell.createWithDataType DataType.String 3 3 "C3"
        ]
    ]
    |> Seq.concat
let dummyFsCellsCollection = FsCellsCollection()
dummyFsCells |> Seq.iter (dummyFsCellsCollection.Add >> ignore)



[<Tests>]
let fsCellsCollectionTests =
    testList "FsCellsCollection" [
        testList "Add | GetCells" [
            testCase "FsCells are present" <| fun _ ->
                let gottenCells = dummyFsCellsCollection.GetCells()
                Expect.containsAll gottenCells dummyFsCells "Cells were not added properly"
        ]
        testList "GetFirstAddress" [
            testCase "First FsAddress is correct" <| fun _ ->
                let gottenFirstAddress = dummyFsCellsCollection.GetFirstAddress() |> fun a -> a.RowNumber, a.ColumnNumber
                let expectedFirstAddress = FsAddress(1, 1) |> fun a -> a.RowNumber, a.ColumnNumber
                Expect.equal gottenFirstAddress expectedFirstAddress "First addresses are not equal"
        ]
        testList "GetLastAddress" [
            testCase "Last FsAddress is correct" <| fun _ ->
                let gottenFirstAddress= dummyFsCellsCollection.GetLastAddress() |> fun a -> a.RowNumber, a.ColumnNumber
                let expectedFirstAddress = FsAddress(3, 3) |> fun a -> a.RowNumber, a.ColumnNumber
                Expect.equal gottenFirstAddress expectedFirstAddress "Last addresses are not equal"
        ]
    ]