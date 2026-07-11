namespace FsSpreadsheet.Js


module JsCell =

    open Fable.Core
    open Fable.Core.JsInterop
    open FsSpreadsheet
    open Fable.ExcelJs

    [<Emit("console.log($0)")>]
    let private log (obj:obj) = jsNative

    [<Emit("(function (v) { const d = new Date(v); return new Date(d.getTime() + d.getTimezoneOffset() * 60000); })($0)")>]
    let private normalizeJsDate (value: obj) : System.DateTime = jsNative

    let writeFromFsCell (fsCell: FsCell) = 
        match fsCell.DataType with
        | Boolean   -> 
            fsCell.ValueAsBool() |> box |> Some
        | Number    -> 
            fsCell.ValueAsFloat() |> box |> Some
        | Date      -> 
            let dt = fsCell.ValueAsDateTime() |> System.DateTimeOffset
            let dt = dt + dt.Offset |> box |> Some
            dt
        | String    -> 
            fsCell.Value |> Some
        | Empty ->
            fsCell.Value |> box |> Some 

    /// <summary>
    /// `worksheetName`, `rowIndex` and `columnIndex` are only used for debugging.
    /// </summary>
    /// <param name="worksheetName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="columnIndex"></param>
    /// <param name="jsCell"></param>
    let readToFsCell worksheetName rowIndex columnIndex (jsCell: Cell) =
        let t = enum<Unions.ValueType>(jsCell.``type``)
        let fsadress = FsAddress.fromString(jsCell.address)
        let createFscell = fun dt v -> FsCell(v,dt,address = fsadress)
        let vTemp = string jsCell.value.Value
        let fscell =
            match t with
            | ValueType.Boolean -> 
                let b = System.Boolean.Parse vTemp
                createFscell DataType.Boolean b
            | ValueType.Number -> float vTemp |> createFscell DataType.Number
            | ValueType.Date -> 
                normalizeJsDate jsCell.value.Value |> createFscell DataType.Date
            | ValueType.String -> vTemp |> createFscell DataType.String
            | ValueType.Formula -> 
                match jsCell.formula with
                | "TRUE()" -> 
                    let b = true
                    createFscell DataType.Boolean b
                | "FALSE()" ->
                    let b = false
                    createFscell DataType.Boolean b
                | anyElse ->
                    //let msg = sprintf "ValueType 'Format' (%s) is not fully implemented in FsSpreadsheet and is handled as string input. In %s: (%i,%i)" anyElse worksheetName rowIndex columnIndex
                    //log msg
                    anyElse |> createFscell DataType.String
            | ValueType.Hyperlink ->
                //log (c.value.Value?text)
                jsCell.value.Value?hyperlink |> createFscell DataType.String
            | anyElse ->
                //let msg = sprintf "ValueType `%A` (%s) is not fully implemented in FsSpreadsheet and is handled as string input. In %s: (%i,%i)" anyElse vTemp worksheetName rowIndex columnIndex
                //log msg
                vTemp |> createFscell DataType.String
        fscell


