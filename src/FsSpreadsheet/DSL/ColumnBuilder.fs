﻿namespace FsSpreadsheet.DSL

open FsSpreadsheet
open FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open Microsoft.FSharp.Linq.RuntimeHelpers
open Microsoft.FSharp.Quotations.Patterns
open Expression


type ColumnBuilder() =

    static member Empty : SheetEntity<ColumnElement list> = SheetEntity.ok []

    // -- Computation Expression methods --> 

    member _.Quote  (quotation: Quotations.Expr<'T>) =
        quotation

    member inline this.Zero() : SheetEntity<ColumnElement list> = SheetEntity.ok []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Column: %s")

    member inline _.Yield(c: ColumnElement) =
        SheetEntity.ok [c]

    member inline _.Yield(cs: ColumnElement list) =
        SheetEntity.ok cs

    member inline _.Yield(c: SheetEntity<ColumnElement>) =
        match c with 
        | Some (c,messages) -> 
            SheetEntity.Some ([c], messages)
        | NoneOptional messages -> 
            NoneOptional messages
        | NoneRequired messages -> 
            NoneRequired messages

    member inline _.Yield(cs: SheetEntity<ColumnElement list>) =
        cs

    member inline _.Yield(c: SheetEntity<CellElement>) =
        match c with 
        | Some ((v,Option.Some i),messages) -> 
            SheetEntity.Some ([ColumnElement.IndexedCell (Row i,v)], messages)
        | Some ((v,None),messages) -> 
            SheetEntity.Some ([ColumnElement.UnindexedCell v], messages)
        | NoneOptional messages -> 
            NoneOptional messages
        | NoneRequired messages -> 
            NoneRequired messages

    member inline _.Yield(c: CellElement) =
        let re = 
            match c with
            | v, Option.Some i -> ColumnElement.IndexedCell (Row i, v)
            | v, None -> ColumnElement.UnindexedCell v
        SheetEntity.ok [re]

    member inline _.Yield(c: SheetEntity<Value>) =
        match c with 
        | Some ((v),messages) -> 
            SheetEntity.Some ([ColumnElement.UnindexedCell v], messages)
        | NoneOptional messages -> 
            NoneOptional messages
        | NoneRequired messages -> 
            NoneRequired messages

    member inline _.Yield(cs: CellElement list) =
        let res = 
            cs 
            |> List.map (function
                | v, Option.Some i -> ColumnElement.IndexedCell (Row i, v)
                | v, None -> ColumnElement.UnindexedCell v
            )
        SheetEntity.ok res

    member inline this.Yield(cs: seq<SheetEntity<CellElement>>) : SheetEntity<ColumnElement list>=
        cs
        |> Seq.map this.Yield
        |> Seq.reduce (fun a b -> this.Combine(a,b))

    member inline this.Yield(n: 'a when 'a :> System.IFormattable) = 
        let v = DataType.InferCellValue n
        SheetEntity.ok [ColumnElement.UnindexedCell v]       

    member inline _.Yield(s : string) = 
        let v = DataType.InferCellValue s
        SheetEntity.ok [ColumnElement.UnindexedCell v]

    member inline this.Yield(n: RequiredSource<'T>) = 
        n

    member inline this.Yield(n: OptionalSource<'T>) = 
        n

    member inline this.YieldFrom(ns: SheetEntity<ColumnElement list> seq) =   
        ns
        |> Seq.fold (fun (state : SheetEntity<ColumnElement list>) we ->
            this.Combine(state,we)

        ) ColumnBuilder.Empty


    member inline this.For(vs : seq<'T>, f : 'T -> SheetEntity<ColumnElement list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    /// Returns the columnelements in SheetEntity container. If the expression does not evaluate, return them es Missing and Optional.
    [<CustomOperation("optional")>] 
    member this.Optional () : SheetEntity<ColumnElement list> =
        //OptionalSource()
        SheetEntity.NoneOptional []

    /// Returns the columnelements in SheetEntity container. If the expression does not evaluate, return them es Missing and Required.
    [<CustomOperation("required"(*,AllowIntoPattern = true*))>] 
    member this.Required (source) (*: SheetEntity<ColumnElement list>*) = 
        RequiredSource(source)
        //SheetEntity.NoneRequired []

    member inline this.Run(children: Expr<SheetEntity<ColumnElement list>>) =
        eval<SheetEntity<ColumnElement list>> children

    member this.Combine(wx1: SheetEntity<ColumnElement list>, wx2: SheetEntity<ColumnElement list>) : SheetEntity<ColumnElement list>=
        match wx1,wx2 with
        // If both contain content, combine the content
        | Some (l1,messages1), Some (l2,messages2) ->
            Some (List.append l1 l2
            ,List.append messages1 messages2)

        // If any of the two is missing and was required, return a missing required
        | _, NoneRequired messages2 ->
            NoneRequired (List.append wx1.Messages messages2)

        | NoneRequired messages1, _ ->
            NoneRequired (List.append messages1 wx2.Messages)

        // If only one of the two is missing and was optional, take the content of the functioning one
        | Some (f1,messages1), NoneOptional messages2 ->
            Some (f1
            ,List.append messages1 messages2)

        | NoneOptional messages1, Some (f2,messages2) ->
            Some (f2
            ,List.append messages1 messages2)

        // If both are missing and were optional, return a missing optional
        | NoneOptional messages1, NoneOptional messages2 ->
            NoneOptional (List.append messages1 messages2)
       
    member this.Combine(wx1: RequiredSource<unit>, wx2: SheetEntity<ColumnElement list>) =
        RequiredSource (wx2)
        
    member this.Combine(wx1: SheetEntity<ColumnElement list>, wx2: RequiredSource<unit>) =
        RequiredSource (wx1)

    member this.Combine(wx1: OptionalSource<unit>, wx2: SheetEntity<ColumnElement list>) =
        OptionalSource wx2

    member this.Combine(wx1: SheetEntity<ColumnElement list>, wx2: OptionalSource<unit>) =
        OptionalSource wx1

    member this.Combine(wx1: RequiredSource<SheetEntity<ColumnElement list>>, wx2: SheetEntity<ColumnElement list>) =
        this.Combine(wx1.Source,wx2) 
        |> RequiredSource

    member this.Combine(wx1: SheetEntity<ColumnElement list>, wx2: RequiredSource<SheetEntity<ColumnElement list>>) =
        this.Combine(wx1,wx2.Source) 
        |> RequiredSource

    member this.Combine(wx1: OptionalSource<SheetEntity<ColumnElement list>>, wx2: SheetEntity<ColumnElement list>) =
        this.Combine(wx1.Source,wx2) 
        |> OptionalSource

    member this.Combine(wx1: SheetEntity<ColumnElement list>, wx2: OptionalSource<SheetEntity<ColumnElement list>>) =
        this.Combine(wx1,wx2.Source) 
        |> OptionalSource

    member inline _.Delay(n: unit -> SheetEntity<ColumnElement list>) = n()


[<AutoOpen>]
module ColumnExtensions =
    type ColumnBuilder with
        [<CompiledName("RunQueryAsColumn")>]
        member this.Run (q: Quotations.Expr<SheetEntity<ColumnElement list>>) = 
            (eval<SheetEntity<ColumnElement list>> q).Value

[<AutoOpen>]
module OptionColumnExtensions =
    type ColumnBuilder with
        [<CompiledName("RunQueryAsOptionalColumn")>]
        member this.Run (q: Quotations.Expr<OptionalSource<SheetEntity<ColumnElement list>>>) = 
            let subExpr = 
                match q with
                | Call(exprOpt, methodInfo, [subExpr]) -> Result.Ok subExpr
                | Call(exprOpt, methodInfo, [ValueWithName(a,b,c);subExpr]) -> Result.Ok subExpr    
                | x ->                     
                    Result.Error $"could not parse option expression as it was not a call: {x}"
            match subExpr with
            | Result.Ok subExpr -> 
                try 
                    match eval<SheetEntity<ColumnElement list>> subExpr with
                    | NoneRequired m -> NoneOptional m
                    | se -> se
                with 
                | err -> NoneOptional([err.Message])  
            | Result.Error err -> NoneOptional([err]) 

[<AutoOpen>]
module RequiredColumnExtensions =
    type ColumnBuilder with
        [<CompiledName("RunQueryAsRequiredColumn")>]
        member this.Run (q: Quotations.Expr<RequiredSource<SheetEntity<ColumnElement list>>>) = 
            let subExpr = 
                match q with
                | Call(exprOpt, methodInfo, [subExpr]) -> Result.Ok subExpr
                | Call(exprOpt, methodInfo, [ValueWithName(a,b,c);subExpr]) -> Result.Ok subExpr    
                | x ->                     
                    Result.Error $"could not parse option expression as it was not a call: {x}"
            match subExpr with
            | Result.Ok subExpr -> 
                try 
                    match eval<SheetEntity<ColumnElement list>> subExpr with
                    | NoneOptional m -> NoneRequired m
                    | se -> se
                with 
                | err -> NoneRequired([err.Message])  
            | Result.Error err -> NoneRequired([err]) 