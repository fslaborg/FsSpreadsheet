﻿module DefaultIO.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Py
open Fable.Openpyxl


let private readFromTestFile (testFile: DefaultTestObject.TestFiles) =
    FsWorkbook.fromXlsxFile(testFile.asRelativePathNode)

let private tests_Read = testList "Read" [

    testCase "ExcelRowCount" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Excel
        Expect.isTrue (wb.GetWorksheetAt(1).Rows.Count > 0) "ExcelRowCount"

    testCase "Excel" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Excel
        Expect.isDefaultTestObject wb
    
    testCase "Libre" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Libre
        Expect.isDefaultTestObject wb
    
    testCase "Pandas" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Pandas
        Expect.isDefaultTestObject wb

    testCase "FableExceljs" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs
        Expect.isDefaultTestObject wb
    
    testCase "ClosedXML" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML
        Expect.isDefaultTestObject wb
    
    testCase "FsSpreadsheetNET" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetNET
        Expect.isDefaultTestObject wb
    
    testCase "FsSpreadsheetJS" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetJS
        Expect.isDefaultTestObject wb
    
]

let private tests_Write = testList "Write" [
    testCase "default" <|  fun _ ->
        let wb = DefaultTestObject.defaultTestObject()
        let p = DefaultTestObject.WriteTestFiles.FsSpreadsheetPY.asRelativePathNode
        do FsWorkbook.toXlsxFile p wb
        let wb_read = FsWorkbook.fromXlsxFile p
        Expect.isDefaultTestObject wb_read
    
]

let main = testList "DefaultIO" [
    tests_Read
    tests_Write
]