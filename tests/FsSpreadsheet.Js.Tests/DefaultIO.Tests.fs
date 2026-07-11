module DefaultIO.Tests

open TestingUtils
open FsSpreadsheet
open FsSpreadsheet.CrossAsync
open FsSpreadsheet.Js

open Fable.Core


let private readFromTestFile (testFile: DefaultTestObject.TestFiles) =      
    FsWorkbook.fromXlsxFile(testFile.asRelativePathNode)
    |> CrossAsync.fromPromise

let private tests_Read = testList "Read" [

    testCaseCrossAsync "Excel" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.Excel
        Expect.isDefaultTestObject wb
    }
    testCaseCrossAsync "Libre" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.Libre
        Expect.isDefaultTestObject wb
    }
    testCaseCrossAsync "Pandas" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.Pandas
        Expect.isDefaultTestObject wb
    }
    testCaseCrossAsync "FableExceljs" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FableExceljs
        Expect.isDefaultTestObject wb
    }
    testCaseCrossAsync "ClosedXML" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.ClosedXML
        Expect.isDefaultTestObject wb
    }
    testCaseCrossAsync "FsSpreadsheetNET" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetNET
        Expect.isDefaultTestObject wb
    }
    testCaseCrossAsync "FsSpreadsheetJS" <| crossAsync {
        let! wb = readFromTestFile DefaultTestObject.TestFiles.FsSpreadsheetJS
        Expect.isDefaultTestObject wb
    }
]

let private tests_Write = testList "Write" [
    testCaseCrossAsync "default" (crossAsync {
        let wb = DefaultTestObject.defaultTestObject()
        let p = DefaultTestObject.WriteTestFiles.FsSpreadsheetJS.asRelativePathNode
        do! CrossAsync.fromPromise(FsWorkbook.toXlsxFile p wb)
        let! wb_read = CrossAsync.fromPromise(FsWorkbook.fromXlsxFile p)
        Expect.isDefaultTestObject wb_read
    })
]

let main = testList "DefaultIO" [
    tests_Read
    tests_Write
]