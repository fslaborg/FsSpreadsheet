﻿module FsWorkbook.Tests
open FsSpreadsheet
open FsSpreadsheet.Net

open TestingUtils

let writeAndReadBytes =
    testList "WriteAndReadBytes" [
        testCase "Empty" (fun () -> 
            let wb = new FsWorkbook()
            let f() = wb.ToXlsxBytes()
            try 
                f() |> ignore
                failwith "Should throw an exception"
            with
            | _ -> ()
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
            let bytes = expected.ToXlsxBytes()
            let actual = FsWorkbook.fromXlsxBytes(bytes)
            Expect.equal (expected.GetWorksheets().Count) (actual.GetWorksheets().Count) "Worksheet count should be equal"
            Expect.equal (expected.GetWorksheetByName(TestObjects.sheet1Name).Name) TestObjects.sheet1Name "excpected sheetname"
            Expect.equal (actual.GetWorksheetByName(TestObjects.sheet1Name).Name) TestObjects.sheet1Name "actual sheetname"
            Expect.workSheetEqual (actual.GetWorksheetByName(TestObjects.sheet1Name)) (expected.GetWorksheetByName(TestObjects.sheet1Name)) "Worksheet did not match"
        )
        testCase "MultipleWorksheets" (fun () -> 
            let wb = new FsWorkbook()
            wb.AddWorksheet(TestObjects.sheet1())
            wb.AddWorksheet(TestObjects.sheet2())
            let bytes = wb.ToXlsxBytes()
            let wb2 = FsWorkbook.fromXlsxBytes(bytes)
            Expect.equal (wb.GetWorksheets().Count) (wb2.GetWorksheets().Count) "Worksheet count should be equal"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet1Name)) (wb2.GetWorksheetByName(TestObjects.sheet1Name)) "First Worksheet did not match"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet2Name)) (wb2.GetWorksheetByName(TestObjects.sheet2Name)) "Second Worksheet did not match"
        )
        testCase "WriteToDisc" (fun () -> 
            let wb = new FsWorkbook()
            let ws1 = TestObjects.sheet1()
            let table = ws1.Table("My Test Table", FsRangeAddress(FsAddress(1,1),FsAddress(3,3)))
            wb.AddWorksheet(ws1)
            wb.AddWorksheet(TestObjects.sheet2())
            System.IO.Directory.CreateDirectory("./TestFiles") |> ignore
            let p = "./TestFiles/WRITE_Roundabout.xlsx"
            wb.ToXlsxFile(p)
            let wb2 = FsWorkbook.fromXlsxFile(p)
            Expect.equal (wb.GetWorksheets().Count) (wb2.GetWorksheets().Count) "Worksheet count should be equal"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet1Name)) (wb2.GetWorksheetByName(TestObjects.sheet1Name)) "First Worksheet did not match"
            Expect.workSheetEqual (wb.GetWorksheetByName(TestObjects.sheet2Name)) (wb2.GetWorksheetByName(TestObjects.sheet2Name)) "Second Worksheet did not match"
        )
        testCase "worksOnFilledTable (issue #100)" <| fun _ ->
            let fsWB = new FsWorkbook()
            let ws = fsWB.InitWorksheet("my awesome worksheet")
            ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("b1")) "my column 1") |> ignore
            ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("c1")) "my column 2") |> ignore
            ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("b2")) 2) |> ignore
            ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("c2")) "row2") |> ignore
            ws.AddTable(FsTable("my_new_table", FsRangeAddress.fromString("b1:c2"))) |> ignore
            fsWB.ToXlsxBytes() |> ignore

        testCase "failsOnEmptyTable (issue #100)" <| fun _ ->
            let fsWB = new FsWorkbook()
            let ws = fsWB.InitWorksheet("my awesome worksheet")
            ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("b1")) "my column 1") |> ignore
            ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("c1")) "my column 2") |> ignore
            ws.AddTable(FsTable("my_new_table", FsRangeAddress.fromString("b1:c1"))) |> ignore
            Expect.throws (fun () -> fsWB.ToXlsxBytes() |> ignore) "no body in table"
    ]

let performance =
    testList "Performace" [
        testCase "ReadBigFile" (fun () -> 
            let sw = Stopwatch()        
            sw.Start()          
            let wb = FsWorkbook.fromXlsxFile(DefaultTestObject.BigFile.asRelativePath)
            sw.Stop()
            let elapsed = sw.Elapsed.Milliseconds
            Expect.isTrue (elapsed < 2000) $"Elapsed time should be less than 2000ms, but was {elapsed}ms"
            Expect.equal (wb.GetWorksheetAt(1).Rows.Count) 153991 "Row count should be 153991"

        )
    ]

let main =
    testList "FsWorkbook" [
        writeAndReadBytes
        performance
    ]