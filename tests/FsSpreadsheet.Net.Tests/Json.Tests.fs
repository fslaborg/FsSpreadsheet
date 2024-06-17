module Json.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Net
open Fable.Pyxpecto

let rows =
    testList "Rows" [

        testCase "Read-Write DefaultTestObject" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToRowsJsonString()
            System.IO.File.WriteAllText(DefaultTestObject.FsSpreadsheetJSON.asRelativePath,s)
            let dto2 = FsWorkbook.fromRowsJsonFile(DefaultTestObject.FsSpreadsheetJSON.asRelativePath)
            Expect.isDefaultTestObject dto2
    ]

let columns =
    testList "Columns" [

        testCase "Read-Write DefaultTestObject" <| fun _ ->
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