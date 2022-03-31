namespace FsSpreadsheet.DSL

open FsSpreadsheet
open FsSpreadsheet.DSL

open Microsoft.FSharp.Quotations
open Expression

type WorkbookBuilder() =

    static member Empty : Missing<WorkbookElement list> = Missing.ok []

    // -- Computation Expression methods --> 

    member inline this.Zero() : Missing<WorkbookElement list> = Missing.ok []

    member this.SignMessages (messages : Message list) : Message list =
        messages
        |> List.map (sprintf "In Workbook: %s")

    member inline _.Yield(c: WorkbookElement) =
        Missing.ok [c]

    member inline _.Yield(cs: WorkbookElement list) =
        Missing.ok cs


    member inline _.Yield(c: Missing<SheetElement list>) =
        match c with 
        | Ok (re,messages) -> 
            Missing.Ok ([WorkbookElement.UnnamedSheet re], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(cs: SheetElement list) =
        Missing.ok [WorkbookElement.UnnamedSheet cs]

    member inline _.Yield(c: Missing<string * SheetElement list>) =
        match c with 
        | Ok ((name,re),messages) -> 
            Missing.Ok ([WorkbookElement.NamedSheet (name,re)], messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member inline _.Yield(cs: string * SheetElement list) =
        Missing.ok [WorkbookElement.NamedSheet (cs)]

    //member inline this.Yield(n: 'a when 'a :> System.IFormattable) = 
    //    let v = DataType.InferCellValue n
    //    Missing.ok [RowElement.UnindexedCell v]

    //member inline _.Yield(s : string) = 
    //    let v = DataType.InferCellValue s
    //    Missing.ok [RowElement.UnindexedCell v]

    member inline this.YieldFrom(ns: (SheetElement list) seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,this.Yield(we))

        ) WorkbookBuilder.Empty

    member inline this.YieldFrom(ns: Missing<SheetElement list> seq) =   
        ns
        |> Seq.fold (fun state we ->
            this.Combine(state,this.Yield(we))

        ) WorkbookBuilder.Empty


    member inline this.For(vs : seq<'T>, f : 'T -> Missing<SheetElement list>) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member inline this.For(vs : seq<'T>, f : 'T -> SheetElement list) =
        vs
        |> Seq.map f
        |> this.YieldFrom

    member inline this.Run(children: Missing<WorkbookElement list>) =
        match children with 
        | Ok (children,messages) -> 
            Missing.Ok (Workbook children, messages)
        | MissingOptional messages -> 
            MissingOptional messages
        | MissingRequired messages -> 
            MissingRequired messages

    member this.Combine(wx1: Missing<WorkbookElement list>, wx2: Missing<WorkbookElement list>) : Missing<WorkbookElement list>=
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
        
    member inline _.Delay(n: unit -> Missing<WorkbookElement list>) = n()
