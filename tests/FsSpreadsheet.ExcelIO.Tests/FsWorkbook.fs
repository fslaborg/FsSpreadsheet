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
        testCase "SingleWorksheet" (fun () -> 
            let wb = new FsWorkbook()
            let ws = TestObjects.sheet1()
            wb.AddWorksheet(ws)
            let bytes = wb.ToBytes()
            let wb2 = FsWorkbook.fromBytes(bytes)
            Expect.equal (wb.GetWorksheets().Count) (wb2.GetWorksheets().Count) "Worksheet count should be equal"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet1Name)) (wb2.GetWorksheetByName(TestObjects.sheet1Name)) "Worksheet did not match"
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