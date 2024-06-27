module FsSpreadsheet.Json.Cell

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let column = "column"

[<Literal>]
let row = "row"

[<Literal>]
let value = "value"

let encodeNoNumber (cell:FsCell) =
    Encode.object [
        value, Value.encode cell.Value
    ]

let encodeRows (cell:FsCell) =
    Encode.object [
        column, Encode.int cell.ColumnNumber
        value, Value.encode cell.Value
    ]

let decodeRows rowNumber : Decoder<FsCell> =
    Decode.object (fun builder ->
        let v,dt = builder.Optional.Field value (Value.decode) |> Option.defaultValue ("", DataType.Empty)
        let c = builder.Optional.Field column Decode.int |> Option.defaultValue 0
        new FsCell(v,dt,FsAddress(Option.defaultValue 0 rowNumber,c))
    )


let encodeCols (cell:FsCell) =
    Encode.object [
        row, Encode.int cell.RowNumber
        value, Value.encode cell.Value
    ]

let decodeCols colNumber : Decoder<FsCell> =
    Decode.object (fun builder ->
        let v,dt = builder.Required.Field value (Value.decode)
        let r = builder.Required.Field row Decode.int
        new FsCell(v,dt,FsAddress(r,colNumber))
    )
