module FsSpreadsheet.Tests
#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto

[<Tests>]
#endif
let all =
    testList "All"
        [
            FsWorkbook.main
            FsWorkSheet.main
            FsTable.main
            FsTableField.main
            FsColumn.main
            FsRow.main
            FsCellsCollection.main
            FsCell.main
            FsAddress.main

            DSL.CellBuilder.main
        ]

[<EntryPoint>]
let main argv = 
    #if FABLE_COMPILER
    Mocha.runTests all
    #else
    Tests.runTestsWithCLIArgs [] argv all
    #endif

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif