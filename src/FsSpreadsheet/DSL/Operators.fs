namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open FsSpreadsheet
open Expression

[<AutoOpen>]
module Operators = 
    
    /// Required value operator
    ///
    /// If expression does fail, returns a missing required value
    let inline (!!) (s : Expr<'a>) : Missing<CellElement> =
        try 
            let value = eval<'a> s |> DataType.InferCellValue
            Missing.ok (value,None)          
        with
        | err -> MissingRequired([err.Message])

    /// Optional value operator
    ///
    /// If expression does fail, returns a missing optional value
    let inline (!?) (s : Expr<'a>) : Missing<CellElement> =
        try 
            let value = eval<'a> s |> DataType.InferCellValue
            Missing.ok (value,None)   
        with
        | err -> MissingOptional([err.Message])