namespace FsSpreadsheet.DSL

open FsSpreadsheet

type Message = string

[<AutoOpen>]
type Missing<'T> =

    | Ok of 'T * Message list
    | MissingOptional of Message list
    | MissingRequired of Message list

    static member ok (v : 'T) : Missing<'T> = Missing.Ok (v,[])

    member this.Value =
        match this with 
        | Ok (f,errs) -> f

    /// Get messages
    member this.Messages =

        match this with 
        | Ok (f,errs) -> errs
        | MissingOptional errs -> errs
        | MissingRequired errs -> errs

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

type SheetElement = 
    | Table of string * TableElement
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
