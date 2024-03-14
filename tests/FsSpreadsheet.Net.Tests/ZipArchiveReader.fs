module ZipArchiveReader

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Net.ZipArchiveReader

let tests_Read = testList "Read" [
    let readFromTestFile (testFile: DefaultTestObject.TestFiles) =
        try 
            let p = testFile.asRelativePath
            FsWorkbook.fromFile(p)
        with
        | err -> 
            printfn "Could not read file from default path: %s" err.Message
            let p = $"{DefaultTestObject.testFolder}/{testFile.asFileName}"
            FsWorkbook.fromFile(p)

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

open FsSpreadsheet.Net

let performanceTest = testList "Performance" [
    testCase "BigFile" <| fun _ ->
        let readF() = FsWorkbook.fromFile(DefaultTestObject.BigFile.asRelativePath)  |> ignore
        let refReadF() = FsWorkbook.fromXlsxFile(DefaultTestObject.BigFile.asRelativePath) |> ignore
        Expect.isFasterThan readF refReadF "ZipArchiveReader should be faster than standard reader"
        //Expect.equal (wb.GetWorksheetAt(1).Rows.Count) 153991 "Row count should be equal"
]


let main = testList "ZipArchiveReader" [
    performanceTest
    tests_Read
]
