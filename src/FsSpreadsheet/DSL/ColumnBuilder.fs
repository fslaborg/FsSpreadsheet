namespace FsSpreadsheet.DSL

open FsSpreadsheet
open FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open Expression

type ColumnBuilder() =

    static member Empty : Missing<ColumnElement list> = Missing.ok []

    // -- Computation Expression methods --> 

    member inline this.Zero() : Missing<ColumnElement list> = Missing.ok []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Column: %s")

    member inline _.Yield(c: ColumnElement) =
        Missing.ok [c]

    member inline _.Yield(cs: ColumnElement list) =
        Missing.ok cs

    member inline _.Yield(cs: Missing<ColumnElement list>) =
        cs

    member inline _.Yield(c: Missing<CellElement>) =
        match c with 
        | Ok ((v,Some i),messages) -> 
            Missing.Ok ([ColumnElement.IndexedCell (Row i,v)], messages)
        | Ok ((v,None),messages) -> 
            Missing.Ok ([ColumnElement.UnindexedCell v], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(c: CellElement) =
        let re = 
            match c with
            | v, Some i -> ColumnElement.IndexedCell (Row i, v)
            | v, None -> ColumnElement.UnindexedCell v
        Missing.ok [re]

    member inline _.Yield(cs: CellElement list) =
        let res = 
            cs 
            |> List.map (function
                | v, Some i -> ColumnElement.IndexedCell (Row i, v)
                | v, None -> ColumnElement.UnindexedCell v
            )
        Missing.ok res

    member inline this.Yield(n: 'a when 'a :> System.IFormattable) = 
        let v = DataType.InferCellValue n
        Missing.ok [ColumnElement.UnindexedCell v]

    member inline _.Yield(s : string) = 
        let v = DataType.InferCellValue s
        Missing.ok [ColumnElement.UnindexedCell v]


    member inline this.YieldFrom(ns: Missing<ColumnElement list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,we)

        ) ColumnBuilder.Empty


    member inline this.For(vs : seq<'T>, f : 'T -> Missing<ColumnElement list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom


    member inline this.Run(children: Missing<ColumnElement list>) =
        children

    member this.Combine(wx1: Missing<ColumnElement list>, wx2: Missing<ColumnElement list>) : Missing<ColumnElement list>=
        match wx1,wx2 with
        // If both contain content, combine the content
        | Ok (l1,messages1), Ok (l2,messages2) ->
            Ok (List.append l1 l2
            ,List.append messages1 messages2)

        // If any of the two is missing and was required, return a missing required
        | _, MissingRequired messages2 ->
            MissingRequired (List.append wx1.Messages messages2)

        | MissingRequired messages1, _ ->
            MissingRequired (List.append messages1 wx2.Messages)

        // If only one of the two is missing and was optional, take the content of the functioning one
        | Ok (f1,messages1), MissingOptional messages2 ->
            Ok (f1
            ,List.append messages1 messages2)

        | MissingOptional messages1, Ok (f2,messages2) ->
            Ok (f2
            ,List.append messages1 messages2)

        // If both are missing and were optional, return a missing optional
        | MissingOptional messages1, MissingOptional messages2 ->
            MissingOptional (List.append messages1 messages2)
        
    member inline _.Delay(n: unit -> Missing<ColumnElement list>) = n()
