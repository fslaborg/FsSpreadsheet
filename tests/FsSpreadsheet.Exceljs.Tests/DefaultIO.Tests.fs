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
    ptestCaseAsync "Libre" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.Libre |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    testCaseAsync "FableExceljs" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    testCaseAsync "ClosedXML" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
    ptestCaseAsync "FsSpreadsheet" <| async {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheet |> Async.AwaitPromise
        Expect.isDefaultTestObject wb
    }
]