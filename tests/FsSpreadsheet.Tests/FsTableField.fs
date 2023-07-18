module FsTableField

open FsSpreadsheet
#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif

let dummyFsRangeAddress = FsRangeAddress("C1:C3")
let dummyFsRangeColumn = FsRangeColumn(dummyFsRangeAddress)
let dummyFsTableField = FsTableField("dummyFsTableField", 0, dummyFsRangeColumn, obj, obj)
let dummyFsCells = [
    FsCell.createWithDataType DataType.String 1 3 "I am the Header!"
    FsCell.createWithDataType DataType.String 2 3 "first data cell"
    FsCell.createWithDataType DataType.String 3 3 "second data cell"
    FsCell.createWithDataType DataType.String 1 4 "Another Header"
    FsCell.createWithDataType DataType.String 2 4 "first data cell in B col"
    FsCell.createWithDataType DataType.String 3 4 "second data cell in B col"
]
let dummyFsCellsCollection = FsCellsCollection()
dummyFsCellsCollection.Add dummyFsCells |> ignore

let main =
    testList "FsTableField" [
        testList "Constructors" [
            testList "unit" [
                let testFsTableField = FsTableField()
                testCase "Index is 0" <| fun _ ->
                    Expect.equal testFsTableField.Index 0 "Index is not 0"
                testCase "Name is empty" <| fun _ ->
                    Expect.equal testFsTableField.Name "" "Name is not empty"
                testCase "Column is null" <| fun _ ->
                    Expect.isNull testFsTableField.Column "Column is not null"
                // TO DO: add testCases for totalsRowLabel & totalsRowFunction as soon as they are implemented
            ]
            testList "name : string" [
                let testFsTableField = FsTableField("testName")
                testCase "Index is 0" <| fun _ ->
                    Expect.equal testFsTableField.Index 0 "Index is not 0"
                testCase "Name is testName" <| fun _ ->
                    Expect.equal testFsTableField.Name "testName" "Name is not testName"
                testCase "Column is null" <| fun _ ->
                    Expect.isNull testFsTableField.Column "Column is not null"
                // TO DO: add testCases for totalsRowLabel & totalsRowFunction as soon as they are implemented
            ]
            testList "name : string, index : int" [
                let testFsTableField = FsTableField("testName", 3)
                testCase "Index is 3" <| fun _ ->
                    Expect.equal testFsTableField.Index 3 "Index is not 3"
                testCase "Name is testName" <| fun _ ->
                    Expect.equal testFsTableField.Name "testName" "Name is not testName"
                testCase "Column is null" <| fun _ ->
                    Expect.isNull testFsTableField.Column "Column is not null"
                // TO DO: add testCases for totalsRowLabel & totalsRowFunction as soon as they are implemented
            ]
            testList "name : string, index : int, column : FsRangeColumn" [
                let testFsRangeAddress = FsRangeAddress("C1:C3")
                let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
                testCase "Index is 3" <| fun _ ->
                    Expect.equal testFsTableField.Index 3 "Index is not 3"
                testCase "Name is testName" <| fun _ ->
                    Expect.equal testFsTableField.Name "testName" "Name is not testName"
                testCase "Column is not null" <| fun _ ->
                    Expect.isNotNull testFsTableField.Column "Column is null"
                testCase "Column range is C1:C3" <| fun _ ->
                    Expect.equal testFsTableField.Column.RangeAddress.Range "C1:C3" "Column range is not C1:C3"
                // TO DO: add testCases for totalsRowLabel & totalsRowFunction as soon as they are implemented
            ]
        ]
        testList "Properties" [
            testList "Index" [
                testList "Setter" [
                    testCase "Sets index correctly" <| fun _ ->
                        let testFsTableField = FsTableField("testName", 3)
                        testFsTableField.Index <- 4
                        Expect.equal testFsTableField.Index 4 "Index is not 4"
                    testCase "Sets column index correctly" <| fun _ ->
                        let testFsRangeAddress = FsRangeAddress("C1:C3")
                        let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                        let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
                        testFsTableField.Index <- 4
                        Expect.equal testFsTableField.Column.RangeAddress.Range "D1:D3" "Range is not D1:D3"
                ]
            ]
            testList "Column" [
                testList "Setter" [
                    let testFsRangeAddress = FsRangeAddress("C1:C3")
                    let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                    let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
                    let testFsRangeAddress = FsRangeAddress "B1:B3"
                    let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                    testFsTableField.Column <- testFsRangeColumn
                    testCase "Sets Column correctly, check Index" <| fun _ ->
                        Expect.equal testFsTableField.Column.Index 2 "ColumnIndex is not 2"
                    testCase "Sets Column correctly, check RangeAddress string" <| fun _ ->
                        Expect.equal testFsTableField.Column.RangeAddress.Range "B1:B3" "ColumnIndex is not 2"
                ]
            ]
        ]
        testList "Methods" [
            testList "HeaderCell" [
                let testFsRangeAddress = FsRangeAddress("C1:C3")
                let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
                let testFsCellsCollection = FsCellsCollection()
                let testFsCells =
                    dummyFsCells
                    |> List.map (
                        fun c -> FsCell.createWithDataType c.DataType c.RowNumber c.ColumnNumber c.Value
                    )
                testFsCellsCollection.Add testFsCells |> ignore
                let headerCell = testFsTableField.HeaderCell(testFsCellsCollection, true)
                testCase "Gets correct header cell" <| fun _ ->
                    Expect.equal headerCell.Value "I am the Header!" "Value is not I am the Header!"
            ]
            testList "SetName" [
                let testFsRangeAddress = FsRangeAddress("C1:C3")
                let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
                let testFsCellsCollection = FsCellsCollection()
                let testFsCells =
                    dummyFsCells
                    |> List.map (
                        fun c -> FsCell.createWithDataType c.DataType c.RowNumber c.ColumnNumber c.Value
                    )
                testFsCellsCollection.Add testFsCells |> ignore
                testCase "Sets fieldname correctly" <| fun _ ->
                    testFsTableField.SetName("testName2", testFsCellsCollection, true)
                    Expect.equal testFsTableField.Name "testName2" "Name is not testName2"
                testCase "Sets first cell's value correctly" <| fun _ ->
                    testFsTableField.SetName("testName2", testFsCellsCollection, true)
                    let headerCell = testFsTableField.HeaderCell(testFsCellsCollection, true)
                    Expect.equal headerCell.Value "testName2" "Value is not testName2"
            ]
            testList "DataCells" [
                let testFsRangeAddress = FsRangeAddress("C1:C3")
                let testFsRangeColumn = FsRangeColumn testFsRangeAddress
                let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
                let testFsCellsCollection = FsCellsCollection()
                let testFsCells =
                    dummyFsCells
                    |> List.map (
                        fun c -> FsCell.createWithDataType c.DataType c.RowNumber c.ColumnNumber c.Value
                    )
                testFsCellsCollection.Add testFsCells |> ignore
                testList "showHeaderRow = true" [
                    testCase "Returns correct data cells" <| fun _ ->
                        let dataCells = testFsTableField.DataCells(testFsCellsCollection)
                        let dataCellsVals = dataCells |> Seq.map (fun c -> c.Value)
                        let col3Cells = testFsCellsCollection.GetCellsInColumn 3
                        let col3CellsNoHeader = col3Cells |> Seq.skip 1
                        let col3CellsNoHeaderVals = col3CellsNoHeader |> Seq.map (fun c -> c.Value)
                        Expect.mySequenceEqual dataCellsVals col3CellsNoHeaderVals "Values of data cells are not equal to values of expected cells"
                ]
                //testCase "Gets correct header cell" <| fun _ ->
                //    Expect.equal headerCell.Value "I am the Header!" "Value is not I am the Header!"
            ]
        ]
    ]