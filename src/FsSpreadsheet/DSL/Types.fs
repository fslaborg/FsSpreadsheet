namespace FsSpreadsheet.DSL

open FsSpreadsheet

type Message = string

[<AutoOpen>]
type SheetEntity<'T> =

    | Some of 'T * Message list
    | NoneOptional of Message list
    | NoneRequired of Message list

    static member ok (v : 'T) : SheetEntity<'T> = SheetEntity.Some (v,[])

    member this.Value =
        match this with 
        | Some (f,errs) -> f

    /// Get messages
    member this.Messages =

        match this with 
        | Some (f,errs) -> errs
        | NoneOptional errs -> errs
        | NoneRequired errs -> errs

type Value = DataType * string

type CellElement = Value * int option

type ColumnIndex = 
    
    | Col of int 

    member self.Index = match self with | Col i -> i

type RowIndex = 
    
    | Row of int

    member self.Index = match self with | Row i -> i

type ColumnElement =
    | IndexedCell of RowIndex * Value
    | UnindexedCell of Value

type RowElement =
    | IndexedCell of ColumnIndex * Value
    | UnindexedCell of Value

type TableElement = 
    | UnindexedRow of RowElement list
    | UnindexedColumn of ColumnElement list

    member this.IsRow =
        match this with 
        | UnindexedRow _ -> true
        | _ -> false

    member this.IsColumn =
        match this with 
        | UnindexedColumn _ -> true
        | _ -> false

type SheetElement = 
    | Table of string * TableElement list
    | IndexedRow of RowIndex * RowElement list
    | UnindexedRow of RowElement list
    | IndexedColumn of ColumnIndex * ColumnElement list
    | UnindexedColumn of ColumnElement list
    | IndexedCell of RowIndex * ColumnIndex * Value
    | UnindexedCell of Value


type WorkbookElement =
    | UnnamedSheet of SheetElement list
    | NamedSheet of string * SheetElement list

type Workbook =
    | Workbook of WorkbookElement list
