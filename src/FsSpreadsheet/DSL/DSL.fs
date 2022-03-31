namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open FsSpreadsheet
open Expression

[<AutoOpen>]
type DSL =
    
    /// Create an xml value from the given value
    static member inline cell (s : string) : CellElement =
        DataType.InferCellValue s, None

    /// Create an xml value from the given value
    static member inline cell<'T when 'T :> System.IFormattable> (v : 'T) : CellElement =
        DataType.InferCellValue v, None

    /// Create an ok xml value from the given value expression if it succedes. Else returns a missing required.
    static member inline cell (s : Quotations.Expr<string>) : Missing<CellElement> =
        try 
            let value = eval<string> s |> DataType.InferCellValue
            Missing.ok (value,None)          
        with
        | err -> MissingRequired([err.Message])

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


