module FsRow


#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif
open FsSpreadsheet


let getDummyWorkSheet() = 
    let worksheet = FsWorksheet("Dummy")
    Resources.dummyFsCells
    |> Seq.concat
    |> Seq.iter (fun c -> worksheet.InsertValueAt(c.Value,c.RowNumber,c.ColumnNumber))
    //worksheet.RescanRows()
    worksheet

let main =
    testList "rowOperations" [
        testList "Prerequisites" [
            let dummyWorkSheet = getDummyWorkSheet()
            testCase "CheckWorksheet" <| fun _ ->
                Expect.equal dummyWorkSheet.CellCollection.Count 9 "Cell count is not correct."
        ]
        testList "RowFromRange" [
            let dummyWorkSheet = getDummyWorkSheet()
            let range = FsRangeAddress("A2:C2")
            let row = FsRow(range, dummyWorkSheet.CellCollection)
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal row.Index 2 "Row index is not correct"
            testCase "CorrectRange" <| fun _ ->
                let expectedRange = FsRangeAddress("A2:C2")
                Expect.equal row.RangeAddress.Range expectedRange.Range "Row index is not correct"
            testCase "CorrectCellCount" <| fun _ ->
                Expect.equal (row.Cells |> Seq.length) 3 "Row length is not correct"
            testCase "RetreiveCorrectCell" <| fun _ ->
                Expect.equal (row.[3].Value) "C2" "Did not retreive correct cell"
            testCase "IsEnumerable" <| fun _ ->
                Expect.equal (row |> Seq.toList |> List.length ) 3 "Did not enumerate correctly"
        ]
        testList "RowFromIndex" [
            let dummyWorkSheet = getDummyWorkSheet()
            let row = FsRow.createAt(2, dummyWorkSheet.CellCollection)
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal row.Index 2 "Row index is not correct"
            testCase "CorrectRange" <| fun _ ->
                let expectedRange = FsRangeAddress("A2:C2")
                Expect.equal row.RangeAddress.Range expectedRange.Range "Row index is not correct"
            testCase "CorrectCellCount" <| fun _ ->
                Expect.equal (row.Cells |> Seq.length) 3 "Row length is not correct"
            testCase "RetreiveCorrectCell" <| fun _ ->
                Expect.equal (row.[3].Value) "C2" "Did not retreive correct cell"
            testCase "IsEnumerable" <| fun _ ->
                Expect.equal (row |> Seq.toList |> List.length ) 3 "Did not enumerate correctly"
        ]
        testList "RowRetrieval" [
            let dummyWorkSheet = getDummyWorkSheet()
            testCase "CorrectRowCount" <| fun _ ->
                Expect.equal (dummyWorkSheet.Rows |> Seq.length) 3 "Row count is not correct"
            testCase "DoesNotFail" <| fun _ ->
                let r = dummyWorkSheet.Row(2)
                Expect.equal 1 1 "Failed when retreiving row"
            testCase "HasRangeAddress" <| fun _ ->
                let r = dummyWorkSheet.Row(2)
                Expect.equal r.RangeAddress.Range "A2:C2" "Row range is not correct"
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal (dummyWorkSheet.Row(2).Index) 2 "Row index is not correct"
        ]
        testList "FromTableRetrieval" [
            
            testCase "CorrectRowCount" <| fun _ ->
                let rows = FsTable.dummyFsTable.GetRows(FsTable.dummyFsCellsCollection)
                Expect.equal (rows |> Seq.length) (FsTable.dummyFsTable.RowCount()) "Column count is not correct"
            testCase "Correct values" <| fun _ ->
                let rows = FsTable.dummyFsTable.GetRows(FsTable.dummyFsCellsCollection)
                let expectedValues = ["Name"; "Age"; "Location"]
                Expect.mySequenceEqual (Seq.item 0 rows |> Seq.map FsCell.getValueAsString) expectedValues "Values are not correct"
        ]
        testList "ToDenseRow" [
            testCase "is correct" (fun _ ->
                let cellsCollWithEmpty = FsCellsCollection()
                cellsCollWithEmpty.Add [
                    FsCell.create 1 1 "Kevin"
                    FsCell.createEmptyWithAdress(FsAddress("B1"))
                    FsCell.create 1 3 "Schneider"
                ]
                let row = FsRow(FsRangeAddress("A1:C1"), cellsCollWithEmpty)
                let actual = FsRow.toDenseRow row |> Seq.map FsCell.getValueAsString |> Seq.toList
                let expected = ["Kevin"; ""; "Schneider"]
                Expect.mySequenceEqual actual expected "Row values differ"
            )
        ]
        testList "Base.Cells" [
            testCase "can be retrieved" <| fun _ ->
                let cellsColl = FsCellsCollection()
                cellsColl.Add (FsCell.create 1 1 "Kevin")
                let row = FsRow(FsRangeAddress("A1:A1"), cellsColl)
                let firstCell = row.Cells |> Seq.head
                Expect.equal (FsCell.getValueAsString firstCell) "Kevin" "Did not retrieve"
        ]
        testList "MinColIndex" [
            testCase "can be retrieved" <| fun _ ->
                let cellsColl = FsCellsCollection()
                cellsColl.Add (FsCell.create 1 1 "Kevin")
                let row = FsRow(FsRangeAddress("A1:A1"), cellsColl)
                let minColIndex = row.MinColIndex
                Expect.equal minColIndex 1 "Incorrect index"
        ]
        testList "MaxColIndex" [
            testCase "can be retrieved" <| fun _ ->
                let cellsColl = FsCellsCollection()
                cellsColl.Add (FsCell.create 1 1 "Kevin")
                let row = FsRow(FsRangeAddress("A1:A1"), cellsColl)
                let maxColIndex = row.MaxColIndex
                Expect.equal maxColIndex 1 "Incorrect index"
        ]
    ]
