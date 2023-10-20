namespace FsSpreadsheet.Exceljs


module JsWorksheet =

    open Fable.Core

    [<Emit("console.log($0)")>]
    let private log (obj:obj) = jsNative

    open FsSpreadsheet
    open Fable.ExcelJs
    open Fable.Core.JsInterop

    let boolConverter (bool:bool) =
        match bool with | true -> "1" | false -> "0"

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
                    c.value <- cell.ValueAsBool() |> boolConverter |> box |> Some
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
            row.eachCell(fun (c, columnIndex) ->
                if c.value.IsSome then
                    let t = enum<Unions.ValueType>(c.``type``)
                    let fsadress = FsAddress(c.address)
                    let createFscell = fun dt v -> FsCell(v,dt,address = fsadress)
                    let vTemp = string c.value.Value
                    let fscell =
                        match t with
                        | ValueType.Boolean -> 
                            let b = System.Boolean.Parse vTemp |> boolConverter
                            createFscell DataType.Boolean b
                        | ValueType.Number -> float vTemp |> createFscell DataType.Number
                        | ValueType.Date -> 
                            let dt = System.DateTime.Parse(vTemp).ToUniversalTime()
                            /// Without this step universal time get changed to local time? Exceljs tests will hit this.
                            ///
                            /// Expected item (from test object): C2 : Sat Oct 14 2023 00:00:00 GMT+0200 (Mitteleuropäische Sommerzeit) | Date
                            ///
                            /// Actual item (created here): C2 : Sat Oct 14 2023 02:00:00 GMT+0200 (Mitteleuropäische Sommerzeit) | Date
                            ///
                            /// But logging hour minute showed that the values were given correctly and needed to be reinitialized.
                            let dt = System.DateTime(dt.Year,dt.Month,dt.Day,dt.Hour,dt.Minute, dt.Second)
                            dt |> createFscell DataType.Date
                        | ValueType.String -> vTemp |> createFscell DataType.String
                        | ValueType.Formula -> 
                            match c.formula with
                            | "TRUE()" -> 
                                let b = true |> boolConverter
                                createFscell DataType.Boolean b
                            | "FALSE()" ->
                                let b = false |> boolConverter
                                createFscell DataType.Boolean b
                            | anyElse ->
                                let msg = sprintf "ValueType 'Format' (%s) is not fully implemented in FsSpreadsheet and is handled as string input. In %s: (%i,%i)" anyElse jsws.name rowIndex columnIndex
                                log msg
                                anyElse |> createFscell DataType.String
                        | ValueType.Hyperlink ->
                            log (c.value.Value?text)
                            vTemp |> createFscell DataType.String
                        | anyElse ->
                            let msg = sprintf "ValueType `%A` (%s) is not fully implemented in FsSpreadsheet and is handled as string input. In %s: (%i,%i)" anyElse vTemp jsws.name rowIndex columnIndex
                            log msg
                            vTemp |> createFscell DataType.String
                    fsws.AddCell(fscell) |> ignore
            )
        )
        for jstableref in jsws.getTables() do
            let table = JsTable.fromJsTable jstableref
            fsws.AddTable table |> ignore
        fsws.RescanRows()
        wb.AddWorksheet(fsws)