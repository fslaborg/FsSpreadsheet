namespace FsSpreadsheet.Exceljs


module JsWorksheet =

    open FsSpreadsheet
    open Fable.ExcelJs

    let addFsWorksheet (wb: Workbook) (fsws:FsWorksheet) : unit =
        let rows = fsws.Rows |> List.map (fun x -> x.Cells |> Seq.map (fun x -> {| adress = x.Address.Address; value = x.Value|}) |> List.ofSeq)
        let ws = wb.addWorksheet(fsws.Name)
        let tables = fsws.Tables |> List.map (fun table -> JsTable.fromFsTable fsws.CellCollection table)
        for table in tables do
            ws.addTable(table)
            |> ignore
        // due to the design of fsspreadsheet this might overwrite some of the stuff from tables, 
        // but as it should be the same, this is only a performance sink.
        for row in rows do
            for cell in row do
                let c = ws.getCell(cell.adress)
                c.value <- Some <| box cell.value

    let addJsWorksheet (wb: FsWorkbook) (jsws: Worksheet) : unit =
        let fsws = FsWorksheet(jsws.name)
        for row in jsws.rows do
            row.eachCell(fun (c, rowIndex) ->
                if c.value.IsSome then
                    let t = enum<Unions.ValueType>(c.``type``)
                    let fsadress = FsAddress(c.address)
                    let createFscell = fun v -> FsCell(v,address = fsadress)
                    let vTemp = string c.value.Value
                    let fscell =
                        match t with
                        | ValueType.Boolean -> System.Boolean.Parse(vTemp) |> createFscell
                        | ValueType.Number -> float vTemp |> createFscell
                        | ValueType.Date -> System.DateTime.Parse(vTemp) |> createFscell
                        | ValueType.String -> vTemp |> createFscell
                        | anyElse -> 
                            printfn "Numbertype '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse
                            vTemp |> createFscell
                    fsws.AddCell(fscell) |> ignore
            )
        for table in jsws.tables do
            let tableRef = table.table.Value.tableRef |> FsRangeAddress
            let table = FsTable(table.name, tableRef,showHeaderRow=table.headerRow)
            fsws.AddTable table
            |> ignore
        wb.AddWorksheet(fsws)