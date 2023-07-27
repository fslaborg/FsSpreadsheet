module Workbook.Tests

open Fable.Mocha
open FsSpreadsheet.Exceljs
open Fable.ExcelJs

let private tests_toFsWorkbook = testList "toFsWorkbook" [
    testCase "empty" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.toFsWorkbook jswb
        Expect.passWithMsg "Convert to fswb"
        let fswsList = fswb.GetWorksheets()
        let jswsList = jswb.worksheets
        Expect.equal fswsList.Length jswsList.Length "Both no worksheet"
]

let main = testList "JsWorkbook<->FsWorkbook" [
    tests_toFsWorkbook
]

