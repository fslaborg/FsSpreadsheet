namespace FsSpreadsheet.Py

#if FABLE_COMPILER_PYTHON || !FABLE_COMPILER

module PyCellType =

    open Fable.Core
    open Fable.Core.JsInterop
    open FsSpreadsheet
    open Fable.Openpyxl

    let fromDataTyoe (t: DataType) = 
        match t with
        | Boolean -> CellType.Boolean
        | Number -> CellType.Float
        | Date -> CellType.DateTime
        | String -> CellType.String
        | anyElse -> 
            let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse         
            printfn "%s" msg
            CellType.String

    let toDataType (t: CellType) =
        match t with
        | CellType.Boolean -> DataType.Boolean
        | CellType.Float -> DataType.Number
        | CellType.Integer -> DataType.Number
        | CellType.DateTime -> DataType.Date
        | CellType.String -> DataType.String
        | anyElse -> 
            let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse         
            printfn "%s" msg
            DataType.String

#endif