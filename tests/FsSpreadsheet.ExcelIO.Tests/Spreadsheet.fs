module Spreadsheet

open Expecto
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml


//let testFilePath = @"C:\Repos\CSBiology\FsSpreadsheet\tests\FsSpreadsheet.ExcelIO.Tests\data\testUnit.xlsx"
let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "data", "testUnit.xlsx")
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.ExcelIO
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let sstpFox = wbpFox.SharedStringTablePart
let sstFox = sstpFox.SharedStringTable
let sstFoxInnerText = sstFox.InnerText


//let testSsdFox2 = Packaging.SpreadsheetDocument.Open(testFilePath, false)

//let testDoc = Spreadsheet.fromFile testFilePath false

//testSsdFox = testDoc
//testSsdFox = testSsdFox2



[<Tests>]
let spreadsheetTests =
    testList "Spreadsheet" [
        // is it even possible to test this?
        //testList "fromFile" [
            //testCase "is equal to testSsd" <| fun _ ->
            //    let ssd = Spreadsheet.fromFile testFilePath false
            //    Expect.equal testDoc ssdFox "Both testFiles differ"
            //    testSsd.Close()
        //]
        testList "tryGetSharedStringTable" [
            let sst = Spreadsheet.tryGetSharedStringTable ssdFox
            testCase "is Some" <| fun _ ->
                Expect.isSome sst "Is None"
            testCase "is equal to testSstFox" <| fun _ ->
                Expect.equal sst.Value sstFox "Differs"
            testCase "inner text is correct" <| fun _ ->
                Expect.equal sst.Value.InnerText sstFoxInnerText "Differs"
        ]
        testList "getWorkbookPart" [
            let wbp = Spreadsheet.getWorkbookPart ssdFox
            testCase "is equal to wbpFox" <| fun _ ->
                Expect.equal wbp wbpFox "Differs"
        ]
    ]









//testSsdFox.Close()