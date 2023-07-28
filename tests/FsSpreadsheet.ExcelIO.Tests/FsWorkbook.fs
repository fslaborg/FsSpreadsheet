module FsWorkbook

open Expecto
open FsSpreadsheet
open FsSpreadsheet.ExcelIO

open TestingUtils

let writeAndReadBytes =
    testList "WriteAndReadBytes" [
        testCase "Empty" (fun () -> 
            let wb = new FsWorkbook()
            let bytes = wb.ToBytes()
            let wb2 = FsWorkbook.fromBytes(bytes)
            Expect.equal (wb.GetWorksheets().Count) (wb2.GetWorksheets().Count) "Worksheet count should be equal"
        )
        testCase "ensure table" (fun () -> 
            let expected = new FsWorkbook()
            let ws = TestObjects.sheet1()
            expected.AddWorksheet(ws)
            Expect.equal (expected.GetWorksheets().Count) 1 "worksheet count"
            let worksheet = expected.GetWorksheets().[0]
            Expect.equal (worksheet.Name) TestObjects.sheet1Name "worksheet name"
            Expect.equal (worksheet.GetCellAt(1,1).Value) "A1" "A1"
            Expect.equal (worksheet.GetCellAt(3,3).Value) "C3" "C3"
            Expect.equal (worksheet.Rows.Count) 3 "rows count"
            Expect.equal (worksheet.Rows.[0].Cells |> Seq.length) 3 "row 0 count"
            Expect.equal (worksheet.Rows.[1].Cells |> Seq.length) 3 "row 1 count"
            Expect.equal (worksheet.Rows.[2].Cells |> Seq.length) 3 "row 2 count"
        )
        testCase "SingleWorksheet" (fun () -> 
            let expected = new FsWorkbook()
            let ws = TestObjects.sheet1()
            expected.AddWorksheet(ws)
            let bytes = expected.ToBytes()
            let actual = FsWorkbook.fromBytes(bytes)
            Expect.equal (expected.GetWorksheets().Count) (actual.GetWorksheets().Count) "Worksheet count should be equal"
            Expect.equal (expected.GetWorksheetByName(TestObjects.sheet1Name).Name) TestObjects.sheet1Name "excpected sheetname"
            Expect.equal (actual.GetWorksheetByName(TestObjects.sheet1Name).Name) TestObjects.sheet1Name "actual sheetname"
            Expect.workSheetEqual (actual.GetWorksheetByName(TestObjects.sheet1Name)) (expected.GetWorksheetByName(TestObjects.sheet1Name)) "Worksheet did not match"
        )
        testCase "MultipleWorksheets" (fun () -> 
            let wb = new FsWorkbook()
            wb.AddWorksheet(TestObjects.sheet1())
            wb.AddWorksheet(TestObjects.sheet2())
            let bytes = wb.ToBytes()
            let wb2 = FsWorkbook.fromBytes(bytes)
            Expect.equal (wb.GetWorksheets().Count) (wb2.GetWorksheets().Count) "Worksheet count should be equal"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet1Name)) (wb2.GetWorksheetByName(TestObjects.sheet1Name)) "First Worksheet did not match"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet2Name)) (wb2.GetWorksheetByName(TestObjects.sheet2Name)) "Second Worksheet did not match"
        )
    ]



[<Tests>]
let tests =
    testList "FsWorkbook" [
        writeAndReadBytes
    ]