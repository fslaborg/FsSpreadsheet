module FsSpreadsheet.Json.Row

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let cells = "cells"

let encode (row:FsRow) =
    Encode.object [
        cells, Encode.seq (row.Cells |> Seq.map Cell.encode)
    ]

let decode : Decoder<FsCell seq> =
    Decode.object (fun builder ->
        builder.Required.Field cells (Decode.seq Cell.decode)
    )