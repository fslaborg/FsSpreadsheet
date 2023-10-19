module DefaultIO

open Expecto
open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.ExcelIO

let tests_Read = testList "Read" [
    let readFromTestFile (testFile: DefaultTestObject.TestFiles) =
        try 
            FsWorkbook.fromXlsxFile(testFile.asRelativePath)
        with
            | _ -> FsWorkbook.fromXlsxFile($"{DefaultTestObject.testFolder}/{testFile.asFileName}")

    testCase "Excel" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Excel
        Expect.isDefaultTestObject wb
    testCase "Libre" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Libre
        Expect.isDefaultTestObject wb
    testCase "FableExceljs" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs
        Expect.isDefaultTestObject wb
    testCase "ClosedXML" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML
        Expect.isDefaultTestObject wb
    testCase "FsSpreadsheet" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheet
        Expect.isDefaultTestObject wb
]

[<Tests>]
let main = testList "DefaultIO" [
    tests_Read
]

