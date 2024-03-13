module FsSpreadsheet.Json.Cell

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let value = "value"

let encode (cell:FsCell) =
    Encode.object [
        value, Value.encode cell.Value
    ]

let decode : Decoder<FsCell> =
    Decode.object (fun builder ->
        let v = builder.Required.Field value (Value.decode)
        new FsCell(v)
    )
