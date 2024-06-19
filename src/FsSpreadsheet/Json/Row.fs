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
        cells, Encode.seq (row.Cells |> Seq.map Cell.encodeRows)
    ]

let decode : Decoder<int option*FsCell seq> =
    Decode.object (fun builder ->
        let n = builder.Optional.Field number Decode.int
        let cs = builder.Optional.Field cells (Decode.seq (Cell.decodeRows n)) |> Option.defaultValue Seq.empty
        n,cs
    )