module FsCellsCollection

open Fable.Pyxpecto
open FsSpreadsheet

let dummyFsCells =
    Resources.dummyFsCells
    |> Seq.concat
let dummyFsCellsCollection = FsCellsCollection()
dummyFsCells |> Seq.iter (dummyFsCellsCollection.Add >> ignore)


let main =
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