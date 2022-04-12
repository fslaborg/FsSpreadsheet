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

    static member Empty : Missing<Value list> = Missing.MissingOptional []

    // -- Computation Expression methods --> 

    member inline this.Zero() : Missing<Value list> = Missing.MissingOptional []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Cell: %s")

    member this.Yield(ro : ReduceOperation) : Missing<Value list> =
        reducer <- ro
        Missing.MissingOptional []

    member inline this.Yield(s : string) : Missing<Value list> =
        Missing.ok [DataType.String,s]

    member inline this.Yield(value: Value) : Missing<Value list> =
        Missing.ok [value]

    member inline this.Yield(value: Missing<Value>) : Missing<Value list> =
        match value with 
        | Ok (v,messages) -> 
            Missing.Ok ([v], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline this.Yield(n: 'a when 'a :> System.IFormattable) = 
        let v = DataType.InferCellValue n
        Missing.ok [v]

    member inline this.Yield(s : string option) : Missing<Value list> =
        match s with
        | Some s -> this.Yield s
        | None -> MissingRequired ["Value is missing"]


    member inline this.Yield(n: 'a option when 'a :> System.IFormattable) = 
        match n with
        | Some s -> this.Yield s
        | None -> MissingRequired ["Value is missing"]


    member inline this.YieldFrom(ns: Missing<Value list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,we)

        ) CellBuilder.Empty

    member inline this.For(vs : seq<'T>, f : 'T -> Missing<Value list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member this.Run(children: Missing<Value list>) : Missing<CellElement> =
        match children with
        | Ok (vals,messages) ->
            let cellElement = reducer.Reduce (vals), None
            Missing.Ok(cellElement, messages)
        | MissingRequired messages -> MissingRequired messages
        | MissingOptional messages -> MissingOptional messages

    member this.Combine(wx1: Missing<Value list>, wx2: Missing<Value list>) : Missing<Value list>=
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
        
    member inline _.Delay(n: unit -> Missing<Value list>) = n()
