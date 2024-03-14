module FsSpreadsheet.Json.Worksheet

open FsSpreadsheet
open Thoth.Json.Core

[<Literal>]
let name = "name"

[<Literal>]
let rows = "rows"

[<Literal>]
let tables = "tables"

let encode (sheet:FsWorksheet) =
    sheet.RescanRows()
    Encode.object [
        name, Encode.string sheet.Name
        if Seq.isEmpty sheet.Tables |> not then        
            tables, Encode.seq (sheet.Tables |> Seq.map Table.encode)
        rows, Encode.seq (sheet.Rows |> Seq.map Row.encode)

    ]

let decode : Decoder<FsWorksheet> =
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