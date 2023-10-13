module FsSpreadsheet.Exceljs.Tests

open Fable.Core.JsInterop
open Fable.Mocha
open TestingUtils

let all =
    testList "All"
        [
            Workbook.Tests.main
            DefaultIO.Tests.tests_Read
        ]

[<EntryPoint>]
let main argv = 
    #if !FABLE_COMPILER
    failwith "The test repo FsSpreadsheet.Exceljs.Tests can only be executed in js environment!"
    #endif
    Mocha.runTests !!all