module FsSpreadsheet.Json.Workbook

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let sheets = "sheets"

let encodeRows noNumbering (wb:FsWorkbook) =
    Encode.object [
        sheets, Encode.seq (wb.GetWorksheets() |> Seq.map (Worksheet.encodeRows noNumbering))
    ]

let decodeRows : Decoder<FsWorkbook> =
    Decode.object (fun builder ->
        let wb = new FsWorkbook()
        let ws = builder.Required.Field sheets (Decode.seq Worksheet.decodeRows)
        wb.AddWorksheets(ws)
        wb
    )

let encodeColumns noNumbering (wb:FsWorkbook) =
    Encode.object [
        sheets, Encode.seq (wb.GetWorksheets() |> Seq.map (Worksheet.encodeColumns noNumbering))
    ]

let decodeColumns : Decoder<FsWorkbook> =
    Decode.object (fun builder ->
        let wb = new FsWorkbook()
        let ws = builder.Required.Field sheets (Decode.seq Worksheet.decodeColumns)
        wb.AddWorksheets(ws)
        wb
    )
