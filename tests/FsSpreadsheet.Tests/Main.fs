module FsSpreadsheet.Tests

open Fable.Pyxpecto



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

// This is possibly the most magic used to make this work. 
// Js and ts cannot use `Async.RunSynchronously`, instead they use `Async.StartAsPromise`.
// Here we need the transpiler not to worry about the output type.
#if !FABLE_COMPILER_JAVASCRIPT && !FABLE_COMPILER_TYPESCRIPT
let (!!) (any: 'a) = any
#endif
#if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
open Fable.Core.JsInterop
#endif

[<EntryPoint>]
let main argv = !!Pyxpecto.runTests [||] all
