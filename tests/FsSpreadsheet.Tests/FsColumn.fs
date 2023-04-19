module FsColumn


open Expecto
open FsSpreadsheet

let getDummyWorkSheet() = 
    let worksheet = FsWorksheet("Dummy")
    Resources.dummyFsCells
    |> Seq.concat
    |> Seq.iter (fun c -> worksheet.InsertValueAt(c.Value,c.RowNumber,c.ColumnNumber))
    worksheet

[<Tests>]
let columnOperations =
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
                Expect.equal (column.Cell(3).Value) "B3" "Did not retreive correct cell"
            testCase "IsEnumerable" <| fun _ ->
                Expect.equal (column |> Seq.toList |> List.length ) 3 "Did not enumerate correctly"
        ]
        testList "ColumnFromIndex" [
            let dummyWorkSheet = getDummyWorkSheet()
            let column = FsColumn(2, dummyWorkSheet.CellCollection)
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal column.Index 2 "Column index is not correct"
            testCase "CorrectRange" <| fun _ ->
                let expectedRange = FsRangeAddress("B1:B3")
                Expect.equal column.RangeAddress.Range expectedRange.Range "Column index is not correct"
            testCase "CorrectCellCount" <| fun _ ->
                Expect.equal (column.Cells |> Seq.length) 3 "Column length is not correct"
            testCase "RetreiveCorrectCell" <| fun _ ->
                Expect.equal (column.Cell(3).Value) "B3" "Did not retreive correct cell"
            testCase "IsEnumerable" <| fun _ ->
                Expect.equal (column |> Seq.toList |> List.length ) 3 "Did not enumerate correctly"
        ]
        testList "ColumnRetrieval" [
            let dummyWorkSheet = getDummyWorkSheet()
            testCase "CorrectColumnCount" <| fun _ ->
                Expect.equal (dummyWorkSheet.Columns |> Seq.length) 3 "Column count is not correct"
            testCase "CorrectIndex" <| fun _ ->
                Expect.equal (dummyWorkSheet.Column(2).Index) 2 "Column index is not correct"
        ]
    ]
