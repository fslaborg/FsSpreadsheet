module Json.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Js
open Fable.Pyxpecto

let rows =
    testList "Rows" [

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