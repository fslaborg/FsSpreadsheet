namespace FsSpreadsheet.DSL

open FsSpreadsheet
open FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open Expression

type SheetBuilder(name : string) =

    static member Empty : Missing<SheetElement list> = Missing.ok []

    // -- Computation Expression methods --> 

    member inline this.Zero() : Missing<SheetElement list> = Missing.ok []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Sheet %s: %s" name)

    member inline _.Yield(c: SheetElement) =
        Missing.ok [c]

    member inline _.Yield(cs: SheetElement list) =
        Missing.ok cs

    member inline _.Yield(c: Missing<RowElement list>) =
        match c with 
        | Ok ((re),messages) -> 
            Missing.Ok ([SheetElement.UnindexedRow re], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(cs: RowElement list) =
        Missing.ok [SheetElement.UnindexedRow cs]

    member inline _.Yield(cs: RowBuilder) =
        Missing.ok [SheetElement.UnindexedRow []]

    member inline _.Yield(c: Missing<ColumnElement list>) =
        match c with 
        | Ok ((re),messages) -> 
            Missing.Ok ([SheetElement.UnindexedColumn re], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(cs: ColumnElement list) =
        Missing.ok [SheetElement.UnindexedColumn cs]

    member inline _.Yield(cs: ColumnBuilder) =
        Missing.ok [SheetElement.UnindexedColumn []]

    member inline this.YieldFrom(ns: (RowElement list) seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,this.Yield(we))

        ) SheetBuilder.Empty

    member inline this.YieldFrom(ns: Missing<RowElement list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,this.Yield(we))

        ) SheetBuilder.Empty

    member inline this.YieldFrom(ns: (ColumnElement list) seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,this.Yield(we))

        ) SheetBuilder.Empty

    member inline this.YieldFrom(ns: Missing<ColumnElement list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,this.Yield(we))

        ) SheetBuilder.Empty

    member inline this.For(vs : seq<'T>, f : 'T -> Missing<RowElement list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member inline this.For(vs : seq<'T>, f : 'T -> RowElement list) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member inline this.For(vs : seq<'T>, f : 'T -> Missing<ColumnElement list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member inline this.For(vs : seq<'T>, f : 'T -> ColumnElement list) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member this.Run(children: Missing<SheetElement list>) =
        match children with 
        | Ok (se,messages) -> 
            Missing.Ok ((name,se), messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages


    member this.Combine(wx1: Missing<SheetElement list>, wx2: Missing<SheetElement list>) : Missing<SheetElement list>=
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
        
    member inline _.Delay(n: unit -> Missing<SheetElement list>) = n()
