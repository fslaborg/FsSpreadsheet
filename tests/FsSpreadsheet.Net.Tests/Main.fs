module FsSpreadsheet.Net.Tests

open Fable.Pyxpecto
open TestingUtils

let all =
    testList "All"
        [
            ZipArchiveReader.main
            Stylesheet.main
            DefaultIO.main
            FsExtension.Tests.main
            Cell.Tests.main
            Sheet.Tests.main
            Workbook.Tests.main
            Spreadsheet.Tests.main          
            Table.Tests.main
            FsWorkbook.Tests.main
            Json.Tests.main
        ]

[<EntryPoint>]
let main argv =
    Pyxpecto.runTests [||] all