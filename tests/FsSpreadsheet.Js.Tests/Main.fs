module FsSpreadsheet.Js.Tests

open Fable.Core.JsInterop
open Fable.Pyxpecto
open TestingUtils

let all =
    testList "All"
        [
            Workbook.Tests.main
            DefaultIO.Tests.main
            Json.Tests.main
        ]


[<EntryPoint>]
let main argv =
    Pyxpecto.runTests [||] all