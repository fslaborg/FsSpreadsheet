module FsTable

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif
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


let main =
    testList "FsTable" [
        testList "AddFields" [
            testList "tableFields : seq FsTableField" [
                let testFsTable = FsTable("testFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))
                testFsTable.AddFields dummyFsTableFields |> ignore
                let testNames, testIndeces = 
                    testFsTable.GetFields(dummyFsCellsCollection)
                    |> Seq.map (fun tf -> tf.Name, tf.Index) 
                    |> Seq.toArray 
                    |> Array.unzip
                let dummyNames, dummyIndeces = 
                    dummyFsTableFields 
                    |> Seq.map (fun tf -> tf.Name, tf.Index) 
                    |> Seq.toArray 
                    |> Array.unzip
                testCase "Names are there" <| fun _ ->
                    Expect.mySequenceEqual testNames dummyNames "Names are not equal" 
                testCase "Indeces are there" <| fun _ ->
                    Expect.mySequenceEqual testIndeces dummyIndeces "Indeces are not equal" 
            ]
        ]
        testList "TryGetHeaderCellOfColumn" [
            testList "cellsCollection : FsCellsCollection, colIndex : int" [
                let testHeaderCell = dummyFsTable.TryGetHeaderCellOfColumnAt(dummyFsCellsCollection, 3)
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testHeaderCell "Is None"
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value.Value actualCell.Value "FsCell is incorrect in value"
            ]
            testList "cellsCollection : FsCellsCollection, column : FsRangeColumn" [
                let testHeaderCell = dummyFsTable.TryGetHeaderCellOfColumn(dummyFsCellsCollection, Seq.item 1 dummyFsRangeColumns)
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testHeaderCell "Is None"
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value.Value actualCell.Value "FsCell is incorrect in value"
            ]
        ]
        testList "TryGetHeaderCellOfTableField" [
            testList "cellsCollection : FsCellsCollection, tableFieldIndex : int" [
                let testFsTable = FsTable("testFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))
                testFsTable.AddFields dummyFsTableFields
                let testHeaderCell = testFsTable.TryGetHeaderCellOfTableFieldAt(dummyFsCellsCollection, 1)
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testHeaderCell "Is None"
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value.Value actualCell.Value "FsCell is incorrect in value"
            ]
        ]
        testList "GetHeaderCellOfTableField" [
            testList "cellsCollection : FsCellsCollection, tableField : FsTableField" [
                let testFsTable = FsTable("testFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))
                testFsTable.AddFields dummyFsTableFields
                let testHeaderCell = testFsTable.GetHeaderCellOfTableField(dummyFsCellsCollection, Seq.item 1 dummyFsTableFields)
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 1
                    Expect.equal testHeaderCell.Value actualCell.Value "FsCell is incorrect in value"
            ]
        ]
        testList "TryGetHeaderCellByFieldName" [
            testList "cellsCollection : FsCellsCollection, fieldName : string" [
                let testFsTable = FsTable("testFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))
                testFsTable.AddFields dummyFsTableFields
                let testHeaderCell = testFsTable.TryGetHeaderCellByFieldName(dummyFsCellsCollection, "Location")
                testCase "Is Some" <| fun _ ->
                    Expect.isSome testHeaderCell "Is None"
                testCase "Has correct value" <| fun _ ->
                    let actualCell = dummyFsCells |> Seq.item 2
                    Expect.equal testHeaderCell.Value.Value actualCell.Value "FsCell is incorrect in value"
            ]
        ]
        testList "GetDataCellsOfColumn" [
            testList "cellsCollection : FsCellsCollection, fieldName : FsTableField" [
                let testFsTable = FsTable("testFsTable", FsRangeAddress(dummyFsCellsCollectionFirstAddress, dummyFsCellsCollectionLastAddress))
                testFsTable.AddFields dummyFsTableFields
                let testDataCells = testFsTable.GetDataCellsOfColumnAt(dummyFsCellsCollection, 2)
                testCase "Is Some" <| fun _ ->
                    Expect.isNonEmpty testDataCells "Seq is empty"
                testCase "Has correct values" <| fun _ ->
                    let minRowNo = dummyFsCells |> Seq.map (fun c -> c.RowNumber) |> Seq.min
                    let actualCells = dummyFsCells |> Seq.filter (fun c -> c.ColumnNumber = 2 && c.RowNumber > minRowNo)
                    Expect.mySequenceEqual (testDataCells |> Seq.map (fun c -> c.Value)) (actualCells |> Seq.map (fun c -> c.Value)) "FsCells are incorrect in value"
            ]
        ]
    ]