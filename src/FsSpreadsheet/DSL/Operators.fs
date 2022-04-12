namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open FsSpreadsheet
open Expression

[<AutoOpen>]
module Operators = 
 
    let inline parseExpression (def : string -> Missing<Value>) (s : Expr<'a>) : Missing<Value> =
        try 
            let value = eval<'a> s |> DataType.InferCellValue
            Missing.ok value         
        with
        | err -> def err.Message

    let inline parseOption (def : string -> Missing<Value>) (s : Option<'a>) : Missing<Value> =
        match s with
        | Some value ->
            DataType.InferCellValue value
            |> Missing.ok  
        | None -> def "Value was missing"
    
    let inline parseResult (def : string -> Missing<Value>) (s : Result<'a,exn>) : Missing<Value> =
        match s with
        | Result.Ok value ->
            DataType.InferCellValue value
            |> Missing.ok  
        | Result.Error exn -> def exn.Message

    let inline parseAny (f : string -> Missing<Value>) (v: 'T) : Missing<Value> =
        match box v with
        | :? Expr<string> as e ->           parseExpression f e
        | :? Expr<int> as e ->              parseExpression f e
        | :? Expr<float> as e ->            parseExpression f e
        | :? Expr<single> as e ->           parseExpression f e
        | :? Expr<byte> as e ->             parseExpression f e
        | :? Expr<System.DateTime> as e ->  parseExpression f e

        | :? Option<string> as o ->             parseOption f o
        | :? Option<int> as o ->                parseOption f o
        | :? Option<float> as o ->              parseOption f o
        | :? Option<single> as o ->             parseOption f o
        | :? Option<byte> as o ->               parseOption f o
        | :? Option<System.DateTime> as o ->    parseOption f o

        | :? Result<string,exn> as r -> parseResult f r
        | :? Result<int,exn> as r -> parseResult f r
        | :? Result<float,exn> as r -> parseResult f r
        | :? Result<single,exn> as r -> parseResult f r
        | :? Result<byte,exn> as r -> parseResult f r
        | :? Result<System.DateTime,exn> as r -> parseResult f r

        | v -> failwith $"Could not parse value {v}. Only string,int,float,single,byte,System.DateTime allowed."

    /// Required value operator
    ///
    /// If expression does fail, returns a missing required value
    let inline (!!) (v : 'T) : Missing<Value> =
        let f = fun s -> MissingRequired([s])
        parseAny f v

    /// Optional value operator
    ///
    /// If expression does fail, returns a missing optional value
    let inline (!?) (v : 'T) : Missing<Value> =
        let f = fun s -> MissingOptional([s])
        parseAny f v 