module DefaultIO

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Net

let tests_Read = testList "Read" [
    let readFromTestFile (testFile: DefaultTestObject.TestFiles) =
        try 
            FsWorkbook.fromXlsxFile(testFile.asRelativePath)
        with
            | _ -> FsWorkbook.fromXlsxFile($"{DefaultTestObject.testFolder}/{testFile.asFileName}")

    testCase "FsCell equality" <| fun _ ->
        let c1 = FsCell(1, DataType.Number, FsAddress("A2"))
        let c2 = FsCell(1, DataType.Number, FsAddress("A2"))
        let isStructEqual = c1.StructurallyEquals(c2)
        Expect.isTrue isStructEqual ""
    testCase "Excel" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Excel
        Expect.isDefaultTestObject wb
    testCase "Libre" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Libre
        Expect.isDefaultTestObject wb
    testCase "Pandas" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.Pandas
        Expect.isDefaultTestObject wb
    testCase "FableExceljs" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs
        Expect.isDefaultTestObject wb
    testCase "ClosedXML" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML
        Expect.isDefaultTestObject wb
    testCase "FsSpreadsheet" <| fun _ ->
        let wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetNET
        wb.GetWorksheets().[0].GetCellAt(5,1) |> fun x -> (x.Value, x.DataType) |> printfn "%A"
        Expect.isDefaultTestObject wb
]

let private tests_Write = testList "Write" [
    testCase "default" <| fun _ ->
        let wb = DefaultTestObject.defaultTestObject()
        let p = DefaultTestObject.WriteTestFiles.FsSpreadsheetNET.asRelativePath
        wb.ToXlsxFile(p)
        let wb_read = FsWorkbook.fromXlsxFile p
        Expect.isDefaultTestObject wb_read

]

let main = testList "DefaultIO" [
    tests_Read
    tests_Write
]

