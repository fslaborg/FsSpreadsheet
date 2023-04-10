module FsTable

open Expecto
open FsSpreadsheet
open FSharpAux


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
dummyFsCellsCollection.Add dummyFsCells
let dummyFsCellsCollectionFirstAddress = dummyFsCellsCollection.GetFirstAddress()
let dummyFsCellsCollectionLastAddress = dummyFsCellsCollection.GetLastAddress()
let dummyFsTable = FsTable("dummyFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress), dummyFsCellsCollection)
let dummyFsTableFields =
    let headerRowIndex = 
        dummyFsCells 
        |> Seq.map (fun c -> c.RowNumber)
        |> Seq.min
    let mutable i = -1
    dummyFsCells
    |> Seq.choose (
        fun fsc -> 
            if fsc.RowNumber = headerRowIndex then
                i <- i + 1
                FsTableField(
                    fsc.Value, 
                    i, 
                    dummyFsRangeColumns |> Seq.find (fun t -> t.RangeAddress.FirstAddress.ColumnNumber = fsc.ColumnNumber), 
                    obj, 
                    obj
                )
                |> Option.Some
            else None
    )
    |> List.ofSeq


[<Tests>]
let fsTableTests =
    testList "FsTable" [
        testList "Item" [
            let testItem = dummyFsTable[5,4]
            testCase "Has correct value" <| fun _ ->
                Expect.equal testItem.Value "Newville" "FsCell is incorrect in value"
            testCase "Throws when outside of range" <| fun _ ->
                Expect.throws (fun _ -> dummyFsTable[1,1] |> ignore) "Did not throw although outside of range"
        ]
        testList "GetSlice" [
            testCase "Throws when outside of range" <| fun _ ->
                Expect.throws (fun _ -> dummyFsTable[10 .. 15,0 .. 1] |> ignore) "Did not throw although outside of range"
            testList "[*,4]" [
                let testSlice = dummyFsTable[*,4] |> List.map FsCell.getValueAs<string>
                let expectedSlice = dummyFsCells |> List.ofSeq |> List.takeNth 3 |> List.map FsCell.getValueAs<string>
                testCase "Returned list has correct values" <| fun _ ->
                    Expect.sequenceEqual testSlice expectedSlice "FsCells are incorrect in value"
            ]
            testList "[3,*]" [
                let testSlice = dummyFsTable[3,*] |> List.map FsCell.getValueAs<string>
                let expectedSlice = dummyFsCells |> List.ofSeq |> List.skip 3 |> List.take 3 |> List.map FsCell.getValueAs<string>
                testCase "Returned list has correct values" <| fun _ ->
                    Expect.sequenceEqual testSlice expectedSlice "FsCells are incorrect in value"
            ]
            testList "[3 to 4,3 to 4]" [
                let testSlice = dummyFsTable[3 .. 4,3 .. 4] |> JaggedList.map FsCell.getValueAs<string> |> List.concat
                let expectedSlice = 
                    dummyFsCells 
                    |> List.ofSeq 
                    |> List.choose (
                        fun c -> 
                            match c.Address.RowNumber, c.Address.ColumnNumber with
                            | x,y when x >= 3 && x <= 4 && y >= 3 && y <= 4 -> (Some << FsCell.getValueAs<string>) c 
                            | _ -> None
                    )
                testCase "Returned list has correct values" <| fun _ ->
                    Expect.sequenceEqual testSlice expectedSlice "FsCells are incorrect in value"
            ]
        ]
        testList "GetFirstRowIndex" [
            testCase "Returns correct index" <| fun _ ->
                Expect.equal (dummyFsTable.GetFirstRowIndex()) 2 "Returns incorrect index"
        ]
        testList "GetLastRowIndex" [
            testCase "Returns correct index" <| fun _ ->
                Expect.equal (dummyFsTable.GetLastRowIndex()) 5 "Returns incorrect index"
        ]
        testList "GetFirstColIndex" [
            testCase "Returns correct index" <| fun _ ->
                Expect.equal (dummyFsTable.GetFirstColIndex()) 2 "Returns incorrect index"
        ]
        testList "GetLastColIndex" [
            testCase "Returns correct index" <| fun _ ->
                Expect.equal (dummyFsTable.GetLastColIndex()) 4 "Returns incorrect index"
        ]
        testList "TryGetHeaderCell" [
            testList "colIndex : int" [
                let testHeaderCell = dummyFsTable.TryGetHeaderCell 3
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testHeaderCell "Is None"
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value.Value actualCell.Value "FsCell is incorrect in value"
            ]
            testList "column : FsRangeColumn" [
                let testHeaderCell = dummyFsTable.TryGetHeaderCell(Seq.item 1 dummyFsRangeColumns)
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testHeaderCell "Is None"
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value.Value actualCell.Value "FsCell is incorrect in value"
            ]
        ]
        testList "GetHeaderCell" [
            testList "colIndex : int" [
                let testHeaderCell = dummyFsTable.GetHeaderCell 3
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value actualCell.Value "FsCell is incorrect in value"
            ]
            testList "column : FsRangeColumn" [
                let testHeaderCell = dummyFsTable.GetHeaderCell(Seq.item 1 dummyFsRangeColumns)
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value actualCell.Value "FsCell is incorrect in value"
            ]
        ]
        testList "TryGetDataCells" [
            testList "colIndex : int" [
                let testDataCells = dummyFsTable.TryGetDataCells 3
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testDataCells "Is None"
                testCase "Returned FsCells have correct values" <| fun _ ->
                    let expectedDataCells = 
                        dummyFsCells 
                        |> List.ofSeq 
                        |> List.choose (
                            fun c -> 
                                match c.Address.RowNumber, c.Address.ColumnNumber with
                                | x,y when x >= 3 && y = 3 -> (Some << FsCell.getValueAs<string>) c 
                                | _ -> None
                        )
                    let actualDataCells = testDataCells.Value |> List.ofSeq |> List.map FsCell.getValueAs<string>
                    Expect.sequenceEqual actualDataCells expectedDataCells "FsCell list is incorrect"
            ]
            testList "column : FsRangeColumn" [
                let testDataCells = dummyFsTable.TryGetDataCells(Seq.head dummyFsRangeColumns)
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testDataCells "Is None"
                testCase "Returned FsCells have correct values" <| fun _ ->
                    let expectedDataCells = 
                        dummyFsCells 
                        |> List.ofSeq 
                        |> List.choose (
                            fun c -> 
                                match c.Address.RowNumber, c.Address.ColumnNumber with
                                | x,y when x >= 3 && y = 2 -> (Some << FsCell.getValueAs<string>) c 
                                | _ -> None
                        )
                    let actualDataCells = testDataCells.Value |> List.ofSeq |> List.map FsCell.getValueAs<string>
                    Expect.sequenceEqual actualDataCells expectedDataCells "FsCell list is incorrect"
            ]
            testList "headerCellValue : string" [
                let testDataCells = dummyFsTable.TryGetDataCells(Seq.item 2 dummyFsCells |> FsCell.getValueAs<string>)
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testDataCells "Is None"
                testCase "Returned FsCells have correct values" <| fun _ ->
                    let expectedDataCells = 
                        dummyFsCells 
                        |> List.ofSeq 
                        |> List.choose (
                            fun c -> 
                                match c.Address.RowNumber, c.Address.ColumnNumber with
                                | x,y when x >= 3 && y = 4 -> (Some << FsCell.getValueAs<string>) c 
                                | _ -> None
                        )
                    let actualDataCells = testDataCells.Value |> List.ofSeq |> List.map FsCell.getValueAs<string>
                    Expect.sequenceEqual actualDataCells expectedDataCells "FsCell list is incorrect"
            ]
        ]
        testList "GetDataCells" [
            testList "colIndex : int" [
                let testDataCells = dummyFsTable.GetDataCells 3
                testCase "Returned FsCells have correct values" <| fun _ ->
                    let expectedDataCells = 
                        dummyFsCells 
                        |> List.ofSeq 
                        |> List.choose (
                            fun c -> 
                                match c.Address.RowNumber, c.Address.ColumnNumber with
                                | x,y when x >= 3 && y = 3 -> (Some << FsCell.getValueAs<string>) c 
                                | _ -> None
                        )
                    let actualDataCells = testDataCells |> List.ofSeq |> List.map FsCell.getValueAs<string>
                    Expect.sequenceEqual actualDataCells expectedDataCells "FsCell list is incorrect"
            ]
            testList "column : FsRangeColumn" [
                let testDataCells = dummyFsTable.GetDataCells(Seq.head dummyFsRangeColumns)
                testCase "Returned FsCells have correct values" <| fun _ ->
                    let expectedDataCells = 
                        dummyFsCells 
                        |> List.ofSeq 
                        |> List.choose (
                            fun c -> 
                                match c.Address.RowNumber, c.Address.ColumnNumber with
                                | x,y when x >= 3 && y = 2 -> (Some << FsCell.getValueAs<string>) c 
                                | _ -> None
                        )
                    let actualDataCells = testDataCells |> List.ofSeq |> List.map FsCell.getValueAs<string>
                    Expect.sequenceEqual actualDataCells expectedDataCells "FsCell list is incorrect"
            ]
            testList "headerCellValue : string" [
                let testDataCells = dummyFsTable.GetDataCells(Seq.item 2 dummyFsCells |> FsCell.getValueAs<string>)
                testCase "Returned FsCells have correct values" <| fun _ ->
                    let expectedDataCells = 
                        dummyFsCells 
                        |> List.ofSeq 
                        |> List.choose (
                            fun c -> 
                                match c.Address.RowNumber, c.Address.ColumnNumber with
                                | x,y when x >= 3 && y = 4 -> (Some << FsCell.getValueAs<string>) c 
                                | _ -> None
                        )
                    let actualDataCells = testDataCells |> List.ofSeq |> List.map FsCell.getValueAs<string>
                    Expect.sequenceEqual actualDataCells expectedDataCells "FsCell list is incorrect"
            ]
        ]
        testList "Copy" [
            let testFsTable = dummyFsTable.Copy()
            let testFsTable2 = dummyFsTable.Copy()
            testCase "Copies correctly" <| fun _ ->
                let testName = testFsTable.Name
                let actualName = dummyFsTable.Name
                Expect.equal actualName testName "Differs in name"
                let testRange = testFsTable.RangeAddress.FirstAddress.Address, testFsTable.RangeAddress.LastAddress.Address
                let actualRange = dummyFsTable.RangeAddress.FirstAddress.Address, dummyFsTable.RangeAddress.LastAddress.Address
                Expect.equal actualRange testRange "Differs in range"
                let testShowHeaderRow = testFsTable.ShowHeaderRow
                let actualShowHeaderRow = dummyFsTable.ShowHeaderRow
                Expect.equal actualShowHeaderRow testShowHeaderRow "Differs in ShowHeaderRow"
                let testFccValues = testFsTable.CellsCollection.GetCells() |> Seq.toList |> List.map FsCell.getValueAs<string>
                let actualFccValues = dummyFsTable.CellsCollection.GetCells() |> Seq.toList |> List.map FsCell.getValueAs<string>
                Expect.sequenceEqual actualFccValues testFccValues "Differs in FsCellsCollection values"
            testCase "Does not change original FsTable" <| fun _ ->
                testFsTable2[2,2].Value <- "new Value"
                Expect.equal dummyFsTable[2,2].Value "Name" "Did change original FsTable in terms of FsCell value"
        ]
    ]