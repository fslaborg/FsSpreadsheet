module FsCell

open Expecto
open FsSpreadsheet
//open FsSpreadsheet.ExcelIO

[<Tests>]
let dataTypeTests =
    testList "DataType" [
        testList "InferCellValue bool = true" [
            let boolValTrue = true
            let resultDtTrue, resultStrTrue = DataType.InferCellValue boolValTrue
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTrue = DataType.Boolean) "is not the expected DataType.Boolean"
            testCase "Correct string" <| fun _ ->
                Expect.isTrue (resultStrTrue = "True") "resulting string is not correct"
        ]
        testList "InferCellValue bool = false" [
            let boolValFalse = false
            let resultDtFalse, resultStrFalse = DataType.InferCellValue boolValFalse
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtFalse = DataType.Boolean) "is not the expected DataType.Boolean"
            testCase "Correct string" <| fun _ ->
                Expect.isTrue (resultStrFalse = "False") "resulting string is not correct"
        ]
        testList "InferCellValue string = \"test\"" [
            let stringValTest = "test"
            let resultDtTest, resultStrTest = DataType.InferCellValue stringValTest
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTest = DataType.String) "is not the expected DataType.String"
            testCase "Correct string" <| fun _ ->
                Expect.isTrue (resultStrTest = "test") "resulting string is not correct"
        ]
        testList "InferCellValue string = \"\"" [
            let stringValTest = ""
            let resultDtTest, resultStrTest = DataType.InferCellValue stringValTest
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTest = DataType.String) "is not the expected DataType.String"
            testCase "Correct string" <| fun _ ->
                Expect.isTrue (resultStrTest = "") "resulting string is not correct"
        ]
        testList "InferCellValue char = 1" [
            let charValTest = '1'
            let resultDtTest, resultChrTest = DataType.InferCellValue charValTest
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTest = DataType.String) "is not the expected DataType.String"
            testCase "Correct string" <| fun _ ->
                Expect.isTrue (resultChrTest = "1") "resulting string is not correct"
        ]
        testList "InferCellValue byte = 255uy" [
            let byteValTest = 255uy
            let resultDtTest, resultBytTest = DataType.InferCellValue byteValTest
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTest = DataType.Number) "is not the expected DataType.Number"
            testCase "Correct string" <| fun _ ->
                Expect.equal "255" resultBytTest "resulting string is not correct"
        ]
        testList "InferCellValue sbyte = -10y" [
            let sbyteValTest = -10y
            let resultDtTest, resultSbyTest = DataType.InferCellValue sbyteValTest
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTest = DataType.Number) "is not the expected DataType.Number"
            testCase "Correct string" <| fun _ ->
                Expect.equal "-10" resultSbyTest "resulting string is not correct"
        ]
        testList "InferCellValue int = 0" [
            let intValTest = 0
            let resultDtTest, resultIntTest = DataType.InferCellValue intValTest
            testCase "Correct DataType" <| fun _ ->
                Expect.isTrue (resultDtTest = DataType.Number) "is not the expected DataType.Number"
            testCase "Correct string" <| fun _ ->
                Expect.equal "0" resultIntTest "resulting string is not correct"
        ]
    ]

[<Tests>]
let fsCellTest =
    testList "FsCell" [
        testList "Constructors" [
        
        ]
    ]