module Json.Tests

open FsSpreadsheet
open FsSpreadsheet.Js
open Fable.Pyxpecto

let defaultTestObject =
    testList "defaultTestObject" [

        testCase "Read-Write DefaultTestObject" <| fun _ ->
            let dto = DefaultTestObject.defaultTestObject()
            let s = dto.ToJsonString()
            let dto2 = FsWorkbook.fromJsonString(s)
            TestingUtils.Expect.isDefaultTestObject dto2
    ]

let main = testList "Json" [
    defaultTestObject
]