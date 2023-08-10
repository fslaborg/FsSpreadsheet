module FsColumn

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
            let range = FsRangeAddress("B1:B3")
            let column = FsColumn(range, dummyWorkSheet.CellCollection)
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal column.Index 2 "Column index is not correct"
            testCase "CorrectRange" <| fun _ ->
                let expectedRange = FsRangeAddress("B1:B3")
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
                let expectedRange = FsRangeAddress("B1:B3")
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
    ]
