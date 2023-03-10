module FsCell

open Expecto
open FsSpreadsheet
//open FsSpreadsheet.ExcelIO

[<Tests>]
let dataTypeTests =
    testList "DataType" [
        testCase "InferCellValue bool = true" <| fun _ ->
            let boolValTrue = true
            let resultDt, resultStr = DataType.InferCellValue boolValTrue
            Expect.isTrue (resultDt = DataType.Boolean) "is not expected bool"
            Expect.isTrue (resultStr = "True") "resulting string is not correct"
        testCase "InferCellValue bool = false" <| fun _ ->
            let boolValFalse = false
            let resultDt, resultStr = DataType.InferCellValue boolValFalse
            Expect.isTrue (resultDt = DataType.Boolean) "is not expected bool"
            Expect.isTrue (resultStr = "False") "resulting string is not correct"
    ]