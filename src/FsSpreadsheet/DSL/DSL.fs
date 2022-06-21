namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open FsSpreadsheet
open Expression

[<AutoOpen>]
type DSL =
    
    /// Create a cell from a value
    static member inline cell = CellBuilder()

    /// Create a row from cells
    static member inline row = RowBuilder()
    
    /// Create a column from cells
    static member inline column = ColumnBuilder()

    /// Create a table from either exclusively rows or exclusively columns. 
    static member inline table name = TableBuilder(name)

    /// Create a sheet from rows, tables and columns
    static member inline sheet name = SheetBuilder(name)

    /// Create a workbook from sheets
    static member inline workbook = WorkbookBuilder()

    /// Transforms any given missing element to an optional.
    static member opt (elem : Missing<'T list>) = 
        match elem with
        | Ok (f,messages) -> elem
        | MissingOptional (messages) -> MissingOptional(messages)
        | MissingRequired (messages) -> MissingOptional(messages)

    /// Transforms any given missing element expression to an optional.
    static member opt (elem : Expr<Missing<'T list>>) = 
        try 
            let elem = eval<Missing<'T list>> elem
            DSL.opt elem
        with
        | err -> 
            MissingOptional([err.Message])

    /// Drops the cell with the given message
    static member dropCell message : Missing<Value> = MissingRequired [message]

    /// Drops the row with the given message
    static member dropRow message : Missing<RowElement> = MissingRequired [message]

    /// Drops the column with the given message
    static member dropColumn message : Missing<ColumnElement> = MissingRequired [message]

    /// Drops the sheet with the given message
    static member dropSheet message : Missing<SheetElement> = MissingRequired [message]

    /// Drops the workbook with the given message
    static member dropWorkbook message : Missing<WorkbookElement> = MissingRequired [message]

