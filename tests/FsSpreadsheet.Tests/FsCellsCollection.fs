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
dummyFsCells |> Seq.iter dummyFsCellsCollection.Add




[<Tests>]
let fsCellsCollectionTests =
    testList "FsCellsCollection" [
        testList "GetSlice" [
            testCase "Returns correct items: 1 .. 3, 2 .. 3" <| fun _ ->
                let res = dummyFsCellsCollection[1 .. 3, 2 .. 3]
                let res1_2 = res |> Seq.item 0 |> Seq.item 1
                Expect.equal res1_2.Value "C1" "Returns incorrect FsCell seqs: i: 1, j: 3"
            testCase "Returns correct items: 1 .. 3, 2" <| fun _ ->
                let res = dummyFsCellsCollection[1 .. 3, 2]
                let res1_2 = res |> Seq.item 0
                Expect.equal res1_2.Value "B1" "Returns incorrect FsCell seqs: i: 1, j: 2"
            testCase "Returns correct items: 2, 2 .. 3" <| fun _ ->
                let res = dummyFsCellsCollection[2, 2 .. 3]
                let res1_2 = res |> Seq.item 1
                Expect.equal res1_2.Value "C2" "Returns incorrect FsCell seqs: i: 1, j: 2"
        ]
        testList "Item" [
            testCase "Returns correct FsCell in terms of value" <| fun _ ->
                Expect.equal dummyFsCellsCollection[2,3].Value "C2" "Returns incorrect FsCell value"
            testCase "Throws when FsCell does not exist: inexistent column" <| fun _ ->
                Expect.throws (fun _ -> dummyFsCellsCollection[2,8].Value |> ignore) "Does not throw when column does not exist"
            testCase "Throws when FsCell does not exist: inexistent row" <| fun _ ->
                Expect.throws (fun _ -> dummyFsCellsCollection[8,3].Value |> ignore) "Does not throw when row does not exist"
        ]
        testList "MaxRowNumber" [
            testCase "Returns correct maximum row index" <| fun _ ->
                Expect.equal dummyFsCellsCollection.MaxRowNumber 3 "Is not the expected row index"
        ]
        testList "MinRowNumber" [
            testCase "Returns correct minimum row index" <| fun _ ->
                Expect.equal dummyFsCellsCollection.MinRowNumber 1 "Is not the expected row index"
        ]
        testList "MaxColNumber" [
            testCase "Returns correct maximum column index" <| fun _ ->
                Expect.equal dummyFsCellsCollection.MaxColumnNumber 3 "Is not the expected column index"
        ]
        testList "MinColNumber" [
            testCase "Returns correct minimum column index" <| fun _ ->
                Expect.equal dummyFsCellsCollection.MinColumnNumber 1 "Is not the expected column index"
        ]
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