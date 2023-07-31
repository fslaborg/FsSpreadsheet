namespace FsSpreadsheet.Exceljs


module JsWorksheet =

    open Fable.Core

    [<Emit("console.log($0)")>]
    let private log (obj:obj) = jsNative

    open FsSpreadsheet
    open Fable.ExcelJs
    open Fable.Core.JsInterop

    let addFsWorksheet (wb: Workbook) (fsws:FsWorksheet) : unit =
        fsws.RescanRows()
        let rows = fsws.Rows |> Seq.map (fun x -> x.Cells)
        let ws = wb.addWorksheet(fsws.Name)
        // due to the design of fsspreadsheet this might overwrite some of the stuff from tables, 
        // but as it should be the same, this is only a performance sink.
        for row in rows do
            for cell in row do
                let c = ws.getCell(cell.Address.Address)
                match cell.DataType with
                | Boolean   -> 
                    c.value <- cell.ValueAsBool() |> box |> Some
                | Number    -> 
                    c.value <- cell.ValueAsFloat() |> box |> Some
                | Date      -> 
                    c.value <- cell.ValueAsDateTime() |> box |> Some
                | String    -> 
                    c.value <- cell.Value |> box |> Some
                | anyElse ->
                    let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse
                    #if FABLE_COMPILER_JAVASCRIPT
                    log msg
                    #else
                    printfn "%s" msg
                    #endif
                    c.value <- cell.Value |> box |> Some 
        let tables = fsws.Tables |> Seq.map (fun table -> JsTable.fromFsTable fsws.CellCollection table)
        for table in tables do
            ws.addTable(table) |> ignore

    let addJsWorksheet (wb: FsWorkbook) (jsws: Worksheet) : unit =
        let fsws = FsWorksheet(jsws.name)
        jsws.eachRow(fun (row, rowIndex) ->
            row.eachCell(fun (c, rowIndex) ->
                if c.value.IsSome then
                    let t = enum<Unions.ValueType>(c.``type``)
                    let fsadress = FsAddress(c.address)
                    let createFscell = fun dt v -> FsCell(v,dt,address = fsadress)
                    let vTemp = string c.value.Value
                    let fscell =
                        match t with
                        | ValueType.Boolean -> System.Boolean.Parse(vTemp) |> createFscell DataType.Boolean
                        | ValueType.Number -> float vTemp |> createFscell DataType.Number
                        | ValueType.Date -> System.DateTime.Parse(vTemp) |> createFscell DataType.Date
                        | ValueType.String -> vTemp |> createFscell DataType.String
                        | anyElse -> 
                            let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse
                            #if FABLE_COMPILER_JAVASCRIPT
                            log msg
                            #else
                            printfn "%s" msg
                            #endif
                            vTemp |> createFscell DataType.String
                    fsws.AddCell(fscell) |> ignore
            )
        )
        for jstableref in jsws.getTables() do
            let table = jstableref.table.Value
            let tableRef = table.tableRef |> FsRangeAddress
            let table = FsTable(table.name, tableRef, table.totalsRow, table.headerRow)
            fsws.AddTable table |> ignore
        fsws.RescanRows()
        wb.AddWorksheet(fsws)