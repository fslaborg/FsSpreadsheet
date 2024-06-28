module FsSpreadsheet.Json.Worksheet

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

let encodeRows noNumbering (sheet:FsWorksheet) =
    
    sheet.RescanRows()
    let jRows = 
        if noNumbering then
            [
                for r = 1 to sheet.MaxRowIndex do
                    [
                        for c = 1 to sheet.MaxColumnIndex do
                            match sheet.CellCollection.TryGetCell(r,c) with
                            | Some cell -> cell
                            | None -> new FsCell("")

                    ]
                    |> Row.encodeNoNumbers
            ]
            |> Encode.seq
        else 
            Encode.seq (sheet.Rows |> Seq.map Row.encode)

    Encode.object [
        name, Encode.string sheet.Name
        if Seq.isEmpty sheet.Tables |> not then        
            tables, Encode.seq (sheet.Tables |> Seq.map Table.encode)
        rows, jRows

    ]

let decodeRows : Decoder<FsWorksheet> =
    Decode.object (fun builder ->
        let mutable rowIndex = 0
        
        let n = builder.Required.Field name Decode.string
        let ts = builder.Optional.Field tables (Decode.seq Table.decode)
        let rs = builder.Optional.Field rows (Decode.seq Row.decode) |> Option.defaultValue Seq.empty
        let sheet = new FsWorksheet(n)
        rs
        |> Seq.iter (fun (rowI,cells) -> 
            let mutable colIndex = 0
            let rowI = 
                match rowI with
                | Some i -> i
                | None -> rowIndex + 1
            rowIndex <- rowI
            let r = sheet.Row(rowI)
            cells 
            |> Seq.iter (fun cell ->
                let colI = 
                    match cell.ColumnNumber with
                    | 0 -> colIndex + 1
                    | i -> i
                colIndex <- colI
                let c = r[colI]
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

let encodeColumns noNumbering (sheet:FsWorksheet) =
    sheet.RescanRows()

    let jColumns = 
        if noNumbering then
            [
                for c = 1 to sheet.MaxColumnIndex do
                    [
                        for r = 1 to sheet.MaxRowIndex do
                            match sheet.CellCollection.TryGetCell(r,c) with
                            | Some cell -> cell
                            | None -> new FsCell("")

                    ]
                    |> Column.encodeNoNumbers
            ]
            |> Encode.seq
        else 
            Encode.seq (sheet.Columns |> Seq.map Column.encode)

    Encode.object [
        name, Encode.string sheet.Name
        if Seq.isEmpty sheet.Tables |> not then        
            tables, Encode.seq (sheet.Tables |> Seq.map Table.encode)
        columns, jColumns
    ]

let decodeColumns : Decoder<FsWorksheet> =
    Decode.object (fun builder ->
        let mutable colIndex = 0
        let n = builder.Required.Field name Decode.string
        let ts = builder.Optional.Field tables (Decode.seq Table.decode)
        let cs = builder.Required.Field columns (Decode.seq Column.decode)
        let sheet = new FsWorksheet(n)
        cs
        |> Seq.iter (fun (colI,cells) -> 
            let mutable rowIndex = 0
            let colI = 
                match colI with
                | Some i -> i
                | None -> colIndex + 1
            colIndex <- colI
            let col = sheet.Column(colI)
            cells 
            |> Seq.iter (fun cell ->    
                let rowI = 
                    match cell.RowNumber with
                    | 0 -> rowIndex + 1
                    | i -> i
                rowIndex <- rowI
                let c = col[rowIndex]
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
