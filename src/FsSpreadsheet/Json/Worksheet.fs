﻿module FsSpreadsheet.Json.Worksheet

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let name = "name"

[<Literal>]
let rows = "rows"

[<Literal>]
let columns = "columns"

[<Literal>]
let tables = "tables"

let encodeRows (sheet:FsWorksheet) =
    sheet.RescanRows()
    Encode.object [
        name, Encode.string sheet.Name
        if Seq.isEmpty sheet.Tables |> not then        
            tables, Encode.seq (sheet.Tables |> Seq.map Table.encode)
        rows, Encode.seq (sheet.Rows |> Seq.map Row.encode)

    ]

let decodeRows : Decoder<FsWorksheet> =
    Decode.object (fun builder ->
        let n = builder.Required.Field name Decode.string
        let ts = builder.Optional.Field tables (Decode.seq Table.decode)
        let rs = builder.Required.Field rows (Decode.seq Row.decode)
        let sheet = new FsWorksheet(n)
        rs
        |> Seq.iter (fun (rowI,cells) -> 
            let r = sheet.Row(rowI)
            cells 
            |> Seq.iter (fun cell ->
                let c = r[cell.ColumnNumber]
                c.Value <- cell.Value
                c.DataType <- cell.DataType
            )
        )
        match ts with
        | Some ts -> 
            ts
            |> Seq.iter (fun t -> 
                sheet.AddTable(t) |> ignore
            )
        | None -> ()
        sheet
    )

let encodeColumns (sheet:FsWorksheet) =
    sheet.RescanRows()
    Encode.object [
        name, Encode.string sheet.Name
        if Seq.isEmpty sheet.Tables |> not then        
            tables, Encode.seq (sheet.Tables |> Seq.map Table.encode)
        columns, Encode.seq (sheet.Columns |> Seq.map Column.encode)
    ]

let decodeColumns : Decoder<FsWorksheet> =
    Decode.object (fun builder ->
        let n = builder.Required.Field name Decode.string
        let ts = builder.Optional.Field tables (Decode.seq Table.decode)
        let cs = builder.Required.Field columns (Decode.seq Column.decode)
        let sheet = new FsWorksheet(n)
        cs
        |> Seq.iter (fun (colI,cells) -> 
            let col = sheet.Column(colI)
            cells 
            |> Seq.iter (fun cell ->        
                let c = col[cell.RowNumber]
                c.Value <- cell.Value
                c.DataType <- cell.DataType
            )
        )
        match ts with
        | Some ts -> 
            ts
            |> Seq.iter (fun t -> 
                sheet.AddTable(t) |> ignore
            )
        | None -> ()
        sheet
    )
