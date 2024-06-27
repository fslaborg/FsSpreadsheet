module Json.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Py
open Fable.Pyxpecto

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
        
        testCase "Read-Write DefaultTestObject" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToRowsJsonString()
            let dto2 = FsWorkbook.fromRowsJsonString(s)
            Expect.isDefaultTestObject dto2
    ]

let columns =
    testList "Columns" [

        testCase "Read-Write DefaultTestObject" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToColumnsJsonString()
            let dto2 = FsWorkbook.fromColumnsJsonString(s)
            Expect.isDefaultTestObject dto2
    ]

let main = testList "Json" [
    rows
    columns
]