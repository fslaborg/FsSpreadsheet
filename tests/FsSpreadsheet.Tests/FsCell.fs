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
    testList "FsCell data" [               
        testList "Data | DataType | Adress" [
            let fscellA1_string  = FsCell.create 1 1 "A1"
            let fscellB1_num     = FsCell.create 1 2 1
            let fscellA2_bool    = FsCell.create 1 2 true
            
            testCase "DataType string" <| fun _ ->
                Expect.equal fscellA1_string.DataType DataType.String "is not the expected DataType.String"
            testCase "DataType number" <| fun _ ->
                Expect.equal fscellB1_num.DataType  DataType.Number "is not the expected DataType.Number"
            testCase "DataType boolean" <| fun _ ->
                Expect.equal fscellA2_bool.DataType DataType.Boolean "is not the expected DataType.Boolean"


            testCase "Value: A1" <| fun _ ->
                Expect.equal fscellA1_string.Value "A1" "resulting value is not A1"
            testCase "Value: 1" <| fun _ ->
                Expect.equal fscellB1_num.Value "1" "resulting value is not 1"
            testCase "Value: True" <| fun _ ->
                Expect.equal fscellA2_bool.Value "True" "resulting value is not True"
        

            testCase "Value as string : A1" <| fun _ ->
                Expect.equal (fscellA1_string.GetValueAs<string>()) "A1" "resulting value is not A1 as string"
            testCase "Value as integer: 1 " <| fun _ ->
                Expect.equal (fscellB1_num.GetValueAs<int>()) 1 "resulting value is not 1 as integer"
            testCase "Value as bool: True" <| fun _ ->
                Expect.equal (fscellA2_bool.GetValueAs<bool>()) true "resulting value is not True as bool"

        
            testCase "RowNumber: 1 " <| fun _ ->
                Expect.equal (fscellB1_num.RowNumber) 1 "resulting value is not A1 as string"
            testCase "ColNumber: 2 " <| fun _ ->
                Expect.equal (fscellB1_num.ColumnNumber) 2 "resulting value is not 1 as integer"

        ]
    ]