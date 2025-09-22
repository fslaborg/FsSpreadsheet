module FsSpreadsheet.Json.Table

open FsSpreadsheet
open Thoth.Json.Core


[<Literal>]
let name = "name"

[<Literal>]
let range = "range"

let encode (sheet:FsTable) =
    Encode.object [
        name, Encode.string sheet.Name
        range, Encode.string sheet.RangeAddress.Range
    ]

let decode : Decoder<FsTable> =
    Decode.object (fun builder ->
        let n = builder.Required.Field name Decode.string
        let r = builder.Required.Field range Decode.string
        FsTable(n, FsRangeAddress.fromString(r))
    )