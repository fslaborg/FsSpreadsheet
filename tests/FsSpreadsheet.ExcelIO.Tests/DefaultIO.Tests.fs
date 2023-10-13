module DefaultIO

open Expecto
open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.ExcelIO

[<Tests>]
let tests_Read = testList "Read" [
    let readFromTestFile (testFile: DefaultTestObject.TestFiles) =
        FsWorkbook.fromXlsxFile(testFile.asRelativePath)

    ftestCase "Excel" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Excel
        Expect.isDefaultTestObject wb
    ptestCase "Libre" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Libre
        Expect.isDefaultTestObject wb
    ftestCase "FableExceljs" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs
        Expect.isDefaultTestObject wb
    ftestCase "ClosedXML" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML
        Expect.isDefaultTestObject wb
    ptestCase "FsSpreadsheet" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheet
        Expect.isDefaultTestObject wb
]

