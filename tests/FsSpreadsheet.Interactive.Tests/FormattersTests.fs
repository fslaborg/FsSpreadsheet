module Formatters

open Expecto
open FsSpreadsheet
open FsSpreadsheet.Interactive

let dummyWorkbook = new FsWorkbook()
let dummyWorksheet1 = FsWorksheet("dummyWorksheet1")
let dummyFsCells =
    [
        FsCell.createWithDataType DataType.String 1 1 "A1"
        FsCell.createWithDataType DataType.String 1 2 "B1"
        FsCell.createWithDataType DataType.String 1 3 "C1"
        FsCell.createWithDataType DataType.String 2 1 "A2"
        FsCell.createWithDataType DataType.String 2 3 "C2"
        FsCell.createWithDataType DataType.String 3 1 "A3"
        FsCell.createWithDataType DataType.String 3 2 "B3"
        FsCell.createWithDataType DataType.String 3 3 "C3"
    ]
dummyWorkbook.AddWorksheet dummyWorksheet1 |> ignore
dummyWorksheet1.AddCells dummyFsCells |> ignore

let expectedWorkbookFormat = """<div><style>.fs-table, .fs-th, .fs-td { border: 1px solid black !important; text-align: left; border-collapse: collapse;}</style><h3>FsWorkbook with 1 worksheets:</h3><table class="fs-table"><thead><tr><th class="fs-th"></th><th class="fs-th">Worksheet</th></tr></thead><tbody><tr><td class="fs-td">1</td><td class="fs-td">dummyWorksheet1</td></tr></tbody></table></div>"""

let expectedWorksheetFormat = """<div><style>.fs-table, .fs-th, .fs-td { border: 1px solid black !important; text-align: left; border-collapse: collapse;}</style><h3>dummyWorksheet1</h3><table class="fs-table"><thead><th class="fs-th"></th><th class="fs-th">A</th><th class="fs-th">B</th><th class="fs-th">C</th></thead><tbody><tr><td class="fs-td">1</td><td class="fs-td">A1</td><td class="fs-td">B1</td><td class="fs-td">C1</td></tr><tr><td class="fs-td">2</td><td class="fs-td">A2</td><td class="fs-td"></td><td class="fs-td">C2</td></tr><tr><td class="fs-td">3</td><td class="fs-td">A3</td><td class="fs-td">B3</td><td class="fs-td">C3</td></tr></tbody></table></div>"""

[<Tests>]
let workbookFormattersTest =
    testList "Formatters.FsWorkbook" [
        testCase "html format" (fun _ ->
            Expect.equal (dummyWorkbook |> Formatters.formatWorkbook) expectedWorkbookFormat "Workbook html format is not correct"
        )
    ]

[<Tests>]
let worksheetFormattersTest =
    testList "Formatters.FsWorksheet" [
        testCase "html format" (fun _ ->
            Expect.equal (dummyWorksheet1 |> Formatters.formatWorksheet) expectedWorksheetFormat "Worksheet html format is not correct"
        )
    ]