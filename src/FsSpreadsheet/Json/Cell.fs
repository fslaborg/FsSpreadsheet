module FsSpreadsheet.Json.Cell

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let column = "column"

[<Literal>]
let value = "value"

let encode (cell:FsCell) =
    Encode.object [
        column, Encode.int cell.ColumnNumber
        value, Value.encode cell.Value
    ]

let decode rowNumber : Decoder<FsCell> =
    Decode.object (fun builder ->
        let v,dt = builder.Required.Field value (Value.decode)
        let c = builder.Required.Field column Decode.int
        new FsCell(v,dt,FsAddress(rowNumber,c))
    )
