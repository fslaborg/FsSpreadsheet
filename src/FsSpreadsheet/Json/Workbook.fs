module FsSpreadsheet.Json.Workbook

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let sheets = "sheets"

let encode (wb:FsWorkbook) =
    Encode.object [
        sheets, Encode.seq (wb.GetWorksheets() |> Seq.map Worksheet.encode)
    ]

let decode : Decoder<FsWorkbook> =
    Decode.object (fun builder ->
        let wb = new FsWorkbook()
        let ws = builder.Required.Field sheets (Decode.seq Worksheet.decode)
        wb.AddWorksheets(ws)
        wb
    )