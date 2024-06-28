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

let encodeNoNumbers (col: FsCell seq) =
    Encode.object [
        cells, Encode.seq (col |> Seq.map Cell.encodeNoNumber)
    ]

let decode : Decoder<int option*FsCell seq> =
    Decode.object (fun builder ->
        let n = builder.Optional.Field number Decode.int
        let cs = builder.Optional.Field cells (Decode.seq (Cell.decodeCols n)) |> Option.defaultValue Seq.empty
        n,cs
    )