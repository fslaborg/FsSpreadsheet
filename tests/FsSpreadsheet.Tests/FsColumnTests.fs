module FsColumn

open Fable.Pyxpecto
open FsSpreadsheet

let getDummyWorkSheet() = 
    let worksheet = FsWorksheet("Dummy")
    Resources.dummyFsCells
    |> Seq.concat
    |> Seq.iter (fun c -> worksheet.InsertValueAt(c.Value,c.RowNumber,c.ColumnNumber))
    worksheet

let main =
    testList "columnOperations" [
        testList "Prerequisites" [
            let dummyWorkSheet = getDummyWorkSheet()
            testCase "CheckWorksheet" <| fun _ ->
                Expect.equal dummyWorkSheet.CellCollection.Count 9 "Cell count is not correct."
        ]
        testList "ColumnFromRange" [
            let dummyWorkSheet = getDummyWorkSheet()
            let range = FsRangeAddress.fromString("B1:B3")
            let column = FsColumn(range, dummyWorkSheet.CellCollection)
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal column.Index 2 "Column index is not correct"
            testCase "CorrectRange" <| fun _ ->
                let expectedRange = FsRangeAddress.fromString("B1:B3")
                Expect.equal column.RangeAddress.Range expectedRange.Range "Column index is not correct"
            testCase "CorrectCellCount" <| fun _ ->
                Expect.equal (column.Cells |> Seq.length) 3 "Column length is not correct"
            testCase "RetreiveCorrectCell" <| fun _ ->
                Expect.equal (column.[3].Value) "B3" "Did not retreive correct cell"
            testCase "IsEnumerable" <| fun _ ->
                Expect.equal (column |> Seq.toList |> List.length ) 3 "Did not enumerate correctly"
        ]
        testList "ColumnFromIndex" [
            let dummyWorkSheet = getDummyWorkSheet()
            let column = FsColumn.createAt(2, dummyWorkSheet.CellCollection)
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal column.Index 2 "Column index is not correct"
            testCase "CorrectRange" <| fun _ ->
                let expectedRange = FsRangeAddress.fromString("B1:B3")
                Expect.equal column.RangeAddress.Range expectedRange.Range "Column index is not correct"
            testCase "CorrectCellCount" <| fun _ ->
                Expect.equal (column.Cells |> Seq.length) 3 "Column length is not correct"
            testCase "RetreiveCorrectCell" <| fun _ ->
                Expect.equal (column.[3].Value) "B3" "Did not retreive correct cell"
            testCase "IsEnumerable" <| fun _ ->
                Expect.equal (column |> Seq.toList |> List.length ) 3 "Did not enumerate correctly"
        ]
        testList "ColumnRetrieval" [
            let dummyWorkSheet = getDummyWorkSheet()
            testCase "CorrectColumnCount" <| fun _ ->
                Expect.equal (dummyWorkSheet.Columns |> Seq.length) 3 "Column count is not correct"
            testCase "DoesNotFail" <| fun _ ->
                let c = dummyWorkSheet.Column(2)
                Expect.equal 1 1 "Failed when retreiving column"
            testCase "HasRangeAddress" <| fun _ ->
                let c = dummyWorkSheet.Column(2)
                Expect.equal c.RangeAddress.Range "B1:B3" "Column range is not correct"
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal (dummyWorkSheet.Column(2).Index) 2 "Column index is not correct"
        ]
        testList "FromTableRetrieval" [
            
            testCase "CorrectColumnCount" <| fun _ ->
                let columns = FsTable.dummyFsTable.GetColumns(FsTable.dummyFsCellsCollection)
                Expect.equal (columns |> Seq.length) (FsTable.dummyFsTable.ColumnCount()) "Column count is not correct"
            testCase "Correct values" <| fun _ ->
                let columns = FsTable.dummyFsTable.GetColumns(FsTable.dummyFsCellsCollection)
                let expectedValues = ["Name";"John Doe";"Jane Doe";"Jack Doe"]
                Expect.mySequenceEqual (Seq.item 0 columns |> Seq.map FsCell.getValueAsString) expectedValues "Values are not correct"
        ]
        testList "ToDenseColumn" [
            testCase "is correct" (fun _ ->
                let cellsCollWithEmpty = FsCellsCollection()
                cellsCollWithEmpty.AddMany [
                    FsCell.create 1 1 "Kevin"
                    FsCell.createEmptyWithAdress(FsAddress.fromString("A2"))
                    FsCell.create 3 1 "Schneider"
                ]
                let column = FsColumn(FsRangeAddress.fromString("A1:A3"), cellsCollWithEmpty)
                let actual = FsColumn.toDenseColumn column |> Seq.map FsCell.getValueAsString |> Seq.toList
                let expected = ["Kevin"; ""; "Schneider"]
                Expect.mySequenceEqual actual expected "Column values differ"
            )
        ]
        testList "Base.Cells" [
            testCase "can be retrieved" <| fun _ ->
                let cellsColl = FsCellsCollection()
                cellsColl.Add (FsCell.create 1 1 "Kevin")
                let column = FsColumn(FsRangeAddress.fromString("A1:A1"), cellsColl)
                let firstCell = column.Cells |> Seq.head
                Expect.equal (FsCell.getValueAsString firstCell) "Kevin" "Did not retrieve"
        ]
        testList "MinColIndex" [
            testCase "can be retrieved" <| fun _ ->
                let cellsColl = FsCellsCollection()
                cellsColl.Add (FsCell.create 1 1 "Kevin")
                let column = FsColumn(FsRangeAddress.fromString("A1:A1"), cellsColl)
                let minRowIndex = column.MinRowIndex
                Expect.equal minRowIndex 1 "Incorrect index"
        ]
        testList "MaxColIndex" [
            testCase "can be retrieved" <| fun _ ->
                let cellsColl = FsCellsCollection()
                cellsColl.Add (FsCell.create 1 1 "Kevin")
                let column = FsColumn(FsRangeAddress.fromString("A1:A1"), cellsColl)
                let maxRowIndex = column.MaxRowIndex
                Expect.equal maxRowIndex 1 "Incorrect index"
        ]
    ]
