namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open FsSpreadsheet
open Expression

[<AutoOpen>]
type DSL =
    
    /// Create an xml value from the given value
    static member inline cell = CellBuilder()

    /// Create an xml element with given name
    static member inline row = RowBuilder()
    
    /// Create an xml element with given name
    static member inline column = ColumnBuilder()

    /// Create an xml element with given name
    static member inline sheet name = SheetBuilder(name)

    /// Create an xml element with given name
    static member inline workbook = WorkbookBuilder()

    /// Transforms any given xml element to an ok.
    static member opt (elem : Missing<'T list>) = 
        match elem with
        | Ok (f,messages) -> elem
        | MissingOptional (messages) -> Ok([],messages)
        | MissingRequired (messages) -> Ok([],messages)

    /// Transforms any given xml element expression to an ok.
    static member opt (elem : Expr<Missing<'T list>>) = 
        try 
            let elem = eval<Missing<'T list>> elem
            DSL.opt elem
        with
        | err -> 
            Ok([],[err.Message])

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

