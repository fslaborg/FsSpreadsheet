module Json.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Net
open Fable.Pyxpecto

let getFilledTestWb() =
    let wb = new FsWorkbook()
    let ws = FsWorkbook.initWorksheet "MySheet" wb
    let r1 = ws.Row(1)
    r1.[1].SetValueAs "A1"
    r1.[2].SetValueAs "B1"
    let r2 = ws.Row(2)
    r2.[1].SetValueAs "A2"
    r2.[2].SetValueAs "B2"
    wb

let rows =
    testList "Rows" [
        testCase "Read Standard" <| fun _ -> 
            // Read object taken from https://spreadsheet.dsl.builders/#_sheets_and_rows
            let s = """{
  "sheets": [
    {
      "name": "Sample",
      "rows": [
        {
          "number": 5,
          "cells": [
            {
              "value": "Line 5"
            }
          ]
        },
        {
        },
        {
          "cells": [
            {
              "value": "Line 7"
            }
          ]
        }
      ]
    }
  ]
}"""
            let wb = FsWorkbook.fromRowsJsonString(s)
            let sheet = Expect.wantSome (wb.TryGetWorksheetByName "Sample") "Sheet Sample"
            Expect.equal sheet.Rows.Count 3 "Row count"

            Expect.isTrue (sheet.ContainsRowAt 5) "Row 5"
            let row1 = sheet.Row(5)
            Expect.isTrue (row1.HasCellAt 1) "Row 5 cell"
            Expect.equal (row1.[1].Value) "Line 5" "Row 5 cell value"

            Expect.isTrue (sheet.ContainsRowAt 6) "Row 6"
            let row1 = sheet.Row(6)
            Expect.hasLength row1.Cells 0 "Row 6 cell count"

            Expect.isTrue (sheet.ContainsRowAt 7) "Row 7"
            let row2 = sheet.Row(7)
            Expect.isTrue (row2.HasCellAt 1) "Row 7 cell"
            Expect.equal (row2.[1].Value) "Line 7" "Row 7 cell value"

        testCase "NoNumber Filled Write" <| fun _ ->
            let wb = getFilledTestWb()
            let expectedString = """{
  "sheets": [
    {
      "name": "MySheet",
      "rows": [
        {
          "cells": [
            {
              "value": "A1"
            },
            {
              "value": "B1"
            }
          ]
        },
        {
          "cells": [
            {
              "value": "A2"
            },
            {
              "value": "B2"
            }
          ]
        }
      ]
    }
  ]
}"""
            let s = wb.ToRowsJsonString(noNumbering = true)
            Expect.stringEqual s expectedString "NoNumber Filled Write-Read"
            
        testCase "NoNumber Filled Write-Read" <| fun _ ->
            let wb = getFilledTestWb()
            let s = wb.ToRowsJsonString(noNumbering = true)
            let wb2 = FsWorkbook.fromRowsJsonString(s)
            Expect.workSheetEqual (wb2.GetWorksheetAt(1)) (wb.GetWorksheetAt(1)) "NoNumber Filled Write-Read"

        testCase "NoNumber DefaultTestObject Write-Read_Success" <| fun _ -> 
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToRowsJsonString(noNumbering = true)
            let dto2 = FsWorkbook.fromRowsJsonString(s)
            ()

        testCase "Write-Read DefaultTestObject" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToRowsJsonString()
            System.IO.File.WriteAllText(DefaultTestObject.FsSpreadsheetJSON.asRelativePath,s)
            let dto2 = FsWorkbook.fromRowsJsonFile(DefaultTestObject.FsSpreadsheetJSON.asRelativePath)
            Expect.isDefaultTestObject dto2
    ]

let columns =
    testList "Columns" [

        testCase "NoNumber Filled Write" <| fun _ ->
            let wb = getFilledTestWb()
            let expectedString = """{
  "sheets": [
    {
      "name": "MySheet",
      "columns": [
        {
          "cells": [
            {
              "value": "A1"
            },
            {
              "value": "A2"
            }
          ]
        },
        {
          "cells": [
            {
              "value": "B1"
            },
            {
              "value": "B2"
            }
          ]
        }
      ]
    }
  ]
}"""
            let s = wb.ToColumnsJsonString(noNumbering = true)
            Expect.stringEqual s expectedString "NoNumber Filled Write-Read"
            
        testCase "NoNumber Filled Write-Read" <| fun _ ->
            let wb = getFilledTestWb()
            let s = wb.ToColumnsJsonString(noNumbering = true)
            let wb2 = FsWorkbook.fromColumnsJsonString(s)
            Expect.workSheetEqual (wb2.GetWorksheetAt(1)) (wb.GetWorksheetAt(1)) "NoNumber Filled Write-Read"

        testCase "NoNumber DefaultTestObject Write-Read_Success" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToColumnsJsonString(noNumbering = true)
            let dto2 = FsWorkbook.fromColumnsJsonString(s)
            ()

        testCase "DefaultTestObject Write-Read" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToColumnsJsonString()
            System.IO.File.WriteAllText(DefaultTestObject.FsSpreadsheetJSON.asRelativePath,s)
            let dto2 = FsWorkbook.fromColumnsJsonFile(DefaultTestObject.FsSpreadsheetJSON.asRelativePath)
            Expect.isDefaultTestObject dto2
    ]

let main = testList "Json" [
    rows
    columns
]