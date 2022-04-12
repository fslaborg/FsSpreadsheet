namespace FsSpreadsheet.DSL

open FsSpreadsheet
open FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open Expression

type RowBuilder() =

    static member Empty : Missing<RowElement list> = Missing.ok []

    // -- Computation Expression methods --> 

    member inline this.Zero() : Missing<RowElement list> = Missing.ok []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Row: %s")

    member inline _.Yield(c: RowElement) =
        Missing.ok [c]

    member inline _.Yield(cs: RowElement list) =
        Missing.ok cs

    member inline _.Yield(c: Missing<RowElement>) =
        match c with 
        | Ok (re,messages) -> 
            Missing.Ok ([re], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(cs: Missing<RowElement list>) =
        cs

    member inline _.Yield(c: Missing<CellElement>) =
        match c with 
        | Ok ((v,Some i),messages) -> 
            Missing.Ok ([RowElement.IndexedCell (Col i,v)], messages)
        | Ok ((v,None),messages) -> 
            Missing.Ok ([RowElement.UnindexedCell v], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(c: CellElement) =
        let re = 
            match c with
            | v, Some i -> RowElement.IndexedCell (Col i, v)
            | v, None -> RowElement.UnindexedCell v
        Missing.ok [re]

    member inline _.Yield(cs: CellElement list) =
        let res = 
            cs 
            |> List.map (function
                | v, Some i -> RowElement.IndexedCell (Col i, v)
                | v, None -> RowElement.UnindexedCell v
            )
        Missing.ok res

    member inline this.Yield(n: 'a when 'a :> System.IFormattable) = 
        let v = DataType.InferCellValue n
        Missing.ok [RowElement.UnindexedCell v]

    member inline _.Yield(s : string) = 
        let v = DataType.InferCellValue s
        Missing.ok [RowElement.UnindexedCell v]


    member inline this.YieldFrom(ns: Missing<RowElement list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,we)

        ) RowBuilder.Empty


    member inline this.For(vs : seq<'T>, f : 'T -> Missing<RowElement list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom


    member inline this.Run(children: Missing<RowElement list>) =
        children

    member this.Combine(wx1: Missing<RowElement list>, wx2: Missing<RowElement list>) : Missing<RowElement list>=
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
        
    member inline _.Delay(n: unit -> Missing<RowElement list>) = n()
