module FsSpreadsheet.Json.Worksheet

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let name = "name"

[<Literal>]
let rows = "rows"

let encode (sheet:FsWorksheet) =
    Encode.object [
        name, Encode.string sheet.Name
        rows, Encode.seq (sheet.Rows |> Seq.map Row.encode)
    ]

let decode : Decoder<FsWorksheet> =
    Decode.object (fun builder ->
        let n = builder.Required.Field name Decode.string
        let rs = builder.Required.Field rows (Decode.seq Row.decode)
        let sheet = new FsWorksheet(name)
        rs
        |> Seq.iteri (fun rowI cells -> 
            let r = sheet.Row(rowI + 1)
            cells 
            |> Seq.iteri (fun coli cell ->
                r[coli + 1].Value <- cell.Value
            )
        )
        sheet
    )