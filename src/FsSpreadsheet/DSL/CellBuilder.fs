namespace FsSpreadsheet.DSL

open FsSpreadsheet
open Microsoft.FSharp.Quotations
open Expression

type ReduceOperation =
    | Concat of char

    member this.Reduce (values : Value list) : Value =
        match this with
        | Concat separator -> 
            DataType.String,
            values 
            |> List.map (snd >> string)
            |> List.reduce (fun a b -> $"{a}{separator}{b}")
            
            

type CellBuilder() =

    let mutable reducer = Concat ','

    static member Empty : SheetEntity<Value list> = SheetEntity.NoneOptional []

    // -- Computation Expression methods --> 

    member inline this.Zero() : SheetEntity<Value list> = SheetEntity.NoneOptional []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Cell: %s")

    member this.Yield(ro : ReduceOperation) : SheetEntity<Value list> =
        reducer <- ro
        SheetEntity.NoneOptional []

    member inline this.Yield(s : string) : SheetEntity<Value list> =
        SheetEntity.ok [DataType.String,s]

    member inline this.Yield(value: Value) : SheetEntity<Value list> =
        SheetEntity.ok [value]

    member inline this.Yield(value: SheetEntity<Value>) : SheetEntity<Value list> =
        match value with 
        | Some (v,messages) -> 
            SheetEntity.Some ([v], messages)
        | NoneOptional messages -> 
            NoneOptional messages
        | NoneRequired messages -> 
            NoneRequired messages

    member inline this.Yield(n: 'a when 'a :> System.IFormattable) = 
        let v = DataType.InferCellValue n
        SheetEntity.ok [v]

    member inline this.Yield(s : string option) : SheetEntity<Value list> =
        match s with
        | Option.Some s -> this.Yield s
        | None -> NoneRequired ["Value is missing"]

    member inline this.Yield(n: 'a option when 'a :> System.IFormattable) = 
        match n with
        | Option.Some s -> this.Yield s
        | None -> NoneRequired ["Value is missing"]


    member inline this.YieldFrom(ns: SheetEntity<Value list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,we)

        ) CellBuilder.Empty

    member inline this.For(vs : seq<'T>, f : 'T -> SheetEntity<Value list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member this.Run(children: SheetEntity<Value list>) : SheetEntity<CellElement> =
        match children with
        | Some (vals,messages) ->
            let cellElement = reducer.Reduce (vals), None
            SheetEntity.Some(cellElement, messages)
        | NoneRequired messages -> NoneRequired messages
        | NoneOptional messages -> NoneOptional messages

    member this.Combine(wx1: SheetEntity<Value list>, wx2: SheetEntity<Value list>) : SheetEntity<Value list>=
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
        
    member inline _.Delay(n: unit -> SheetEntity<Value list>) = n()
