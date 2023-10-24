module DefaultIO.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Exceljs
open Fable.Core

let tests_Read = testList "Read" [
    let readFromTestFile (testFile: DefaultTestObject.TestFiles) =
        FsWorkbook.fromXlsxFile(testFile.asRelativePathNode)

    testCaseAsync "Excel" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.Excel |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    testCaseAsync "Libre" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.Libre |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    testCaseAsync "FableExceljs" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    ptestCaseAsync "ClosedXML" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    testCaseAsync "FsSpreadsheetNET" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetNET |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    testCaseAsync "FsSpreadsheetJS" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetJS |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
]