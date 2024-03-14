module DefaultIO.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.Js
open Fable.Core


let private readFromTestFile (testFile: DefaultTestObject.TestFiles) =
    FsWorkbook.fromXlsxFile(testFile.asRelativePathNode)

let private tests_Read = testList "Read" [

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

let private tests_Write = testList "Write" [
    testCaseAsync "default" (Async.AwaitPromise <|  promise {
        let wb = DefaultTestObject.defaultTestObject()
        let p = DefaultTestObject.WriteTestFiles.FsSpreadsheetJS.asRelativePathNode
        do! FsWorkbook.toXlsxFile p wb
        let! wb_read = FsWorkbook.fromXlsxFile p
        Expect.isDefaultTestObject wb_read
    })
]

let main = testList "DefaultIO" [
    tests_Read
    tests_Write
]