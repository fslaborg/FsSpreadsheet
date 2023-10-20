module TestingUtils

open FsSpreadsheet
#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif

[<RequireQualifiedAccess>]
module Utils = 

    let extractWords (json:string) = 
        json.Split([|'{';'}';'[';']';',';':'|])
        |> Array.map (fun s -> s.Trim())
        |> Array.filter ((<>) "")

    let wordFrequency (json:string) = 
        json
        |> extractWords
        |> Array.countBy id
        |> Array.sortBy fst

    let inline firstDiff s1 s2 =
      let s1 = Seq.append (Seq.map Some s1) (Seq.initInfinite (fun _ -> None))
      let s2 = Seq.append (Seq.map Some s2) (Seq.initInfinite (fun _ -> None))
      Seq.mapi2 (fun i s p -> i,s,p) s1 s2
      |> Seq.find (function |_,Some s,Some p when s=p -> false |_-> true)

/// Fable compatible Expecto/Mocha unification
module Expect =

    /// Expects the `actual` sequence to equal the `expected` one.
    let inline private _sequenceEqual message (comparison: int * 'a option * 'a option) =
        match comparison with
        | _,None,None -> ()
        | i,Some a, Some e ->
            let msg = 
                sprintf "%s. Sequence does not match at position %i.\n" message i
                 + sprintf "Expected item: %A\n" e
                 + sprintf "Actual item: %A\n" a
            failwith msg
        | i,None,Some e ->
            let msg =
                sprintf "%s. Sequence actual shorter than expected, at pos %i for expected item: \n" message i
                + sprintf "%A" e
            failwith msg
        | i,Some a,None ->
            let msg =
                sprintf "%s. Sequence actual longer than expected, at pos %i found item: \n" message i
                + sprintf "%A" a
            failwith msg

    let inline sequenceEqual actual expected message =
        let comp = Utils.firstDiff actual expected
        _sequenceEqual message comp
        
    let cellSequenceEquals (actual: FsCell seq) (expected: FsCell seq) message =
        let cellDiff (s1: FsCell seq) (s2: FsCell seq) =
            let s1 = Seq.append (Seq.map Some s1) (Seq.initInfinite (fun _ -> None))
            let s2 = Seq.append (Seq.map Some s2) (Seq.initInfinite (fun _ -> None))
            Seq.mapi2 (fun i s p -> i,s,p) s1 s2
            |> Seq.find (function |_,Some s,Some p when s.StructurallyEquals(p) -> false |_-> true)
        let comp = cellDiff actual expected
        _sequenceEqual message comp

    let columnsEqual (actual : FsCell seq seq) (expected : FsCell seq seq) message =     
        Seq.iteri2 (fun i s1 s2 ->
            cellSequenceEquals s1 s2 $"{message}. Columns do not match in row {i}."
        ) actual expected

    let workSheetEqual (actual : FsWorksheet) (expected : FsWorksheet) message =
        let f (ws : FsWorksheet) = 
            ws.RescanRows()
            ws.Rows |> Seq.map (fun r -> r.Cells) 
        if actual.Name <> expected.Name then
            failwithf $"{message}. Worksheet names do not match. Expected {expected.Name} but got {actual.Name}"
        columnsEqual (f actual) (f expected) $"{message}. Worksheet does not match"

    let isDefaultTestObject (wb: FsWorkbook) = 
        let worksheets = wb.GetWorksheets()
        for ws in worksheets do
            let isTable, expectedRows = Expect.wantSome (DefaultTestObject.valueMap |> Map.tryFind ws.Name) $"ExpectError: Unable to get info for worksheet: {ws.Name}"
            match isTable with
            | Some expectedTableName -> 
                let actualTable = Expect.wantSome (ws.Tables |> Seq.tryFind (fun t -> t.Name = expectedTableName)) $"ExpectError: Unable to get info for worksheet->table: {ws.Name}->{expectedTableName}"
                let headerRow = Expect.wantSome (actualTable.TryGetHeaderRow(ws.CellCollection)) $"ExpectError: ShowHeaderRow is false for worksheet->table: {ws.Name}->{expectedTableName}"
                let actualRows = actualTable.GetRows(ws.CellCollection) |> Seq.tail //Seq.tail skips HeaderRow
                cellSequenceEquals headerRow expectedRows.[0] $"ExpectError: HeaderRow is not equal for worksheet->table: {ws.Name}->{expectedTableName}"
                for actualRow, expectedRow in Seq.zip actualRows expectedRows.[1..] do
                    cellSequenceEquals actualRow expectedRow $"ExpectError: Table body rows are not equal for worksheet->table: {ws.Name}->{expectedTableName}"
            | None ->
                let actualRows = ws.Rows
                for actualRow, expectedRow in Seq.zip actualRows expectedRows do
                    cellSequenceEquals actualRow expectedRow $"ExpectError: Worksheet rows are not equal for worksheet: {ws.Name}"

    let inline equal actual expected message = Expect.equal actual expected message
    let notEqual actual expected message = Expect.notEqual actual expected message

    let isNull actual message = Expect.isNull actual message 
    let isNotNull actual message = Expect.isNotNull actual message 

    let isSome actual message = Expect.isSome actual message 
    let isNone actual message = Expect.isNone actual message 
    let wantSome actual message = Expect.wantSome actual message 

    let isEmpty actual message = Expect.isEmpty actual message 
    let hasLength actual expectedLength message = Expect.hasLength actual expectedLength message

    let isTrue actual message = Expect.isTrue actual message 
    let isFalse actual message = Expect.isFalse actual message 

    let wantError actual message = Expect.wantError actual message 
    let wantOk actual message = Expect.wantOk actual message 
    let isOk actual message = Expect.isOk actual message 
    let isError actual message = Expect.isError actual message 

    let throws actual message = Expect.throws actual message
    let throwsC actual message = Expect.throwsC actual message 

    let exists actual asserter message = Expect.exists actual asserter message 
    let containsAll actual expected message = Expect.containsAll actual expected message

    let passWithMsg (message: string) = equal true true message

/// Fable compatible Expecto/Mocha unification
[<AutoOpen>]
module Test =

    let test = test
    let testAsync = testAsync
    let testSequenced = testSequenced

    let testCase = testCase
    let ptestCase = ptestCase
    let ftestCase = ftestCase
    let testCaseAsync = testCaseAsync
    let ptestCaseAsync = ptestCaseAsync
    let ftestCaseAsync = ftestCaseAsync


    let testList = testList