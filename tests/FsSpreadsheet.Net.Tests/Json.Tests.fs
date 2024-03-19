module Json.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Net
open Fable.Pyxpecto

let defaultTestObject =
    testList "defaultTestObject" [

        testCase "Read-Write DefaultTestObject" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToJsonString()
            System.IO.File.WriteAllText(DefaultTestObject.FsSpreadsheetJSON.asRelativePath,s)
            let dto2 = FsWorkbook.fromJsonString(s)
            TestingUtils.Expect.isDefaultTestObject dto2
    ]

let main = testList "Json" [
    defaultTestObject
]