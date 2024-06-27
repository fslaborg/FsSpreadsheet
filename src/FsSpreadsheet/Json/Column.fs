module FsSpreadsheet.Json.Column

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let cells = "cells"

[<Literal>]
let number = "number"

let encode (col:FsColumn) =
    Encode.object [
        number, Encode.int col.Index
        cells, Encode.seq (col.Cells |> Seq.map Cell.encodeCols)
    ]

let decode : Decoder<int*FsCell seq> =
    Decode.object (fun builder ->
        let n = builder.Required.Field number Decode.int
        let cs = builder.Required.Field cells (Decode.seq (Cell.decodeCols n))
        n,cs
    )