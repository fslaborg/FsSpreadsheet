module FsSpreadsheet.Json.Row

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let cells = "cells"

[<Literal>]
let number = "number"

let encode (row:FsRow) =
    Encode.object [
        number, Encode.int row.Index
        cells, Encode.seq (row.Cells |> Seq.map Cell.encode)
    ]

let decode : Decoder<int*FsCell seq> =
    Decode.object (fun builder ->
        let n = builder.Required.Field number Decode.int
        let cs = builder.Required.Field cells (Decode.seq (Cell.decode n))
        n,cs
    )