module TestUtils.Utils

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

    let firstDiff s1 s2 =
      let s1 = Seq.append (Seq.map Some s1) (Seq.initInfinite (fun _ -> None))
      let s2 = Seq.append (Seq.map Some s2) (Seq.initInfinite (fun _ -> None))
      Seq.mapi2 (fun i s p -> i,s,p) s1 s2
      |> Seq.find (function |_,Some s,Some p when s=p -> false |_-> true)

/// Fable compatible Expecto/Mocha unification
module Expect =

    /// Expects the `actual` sequence to equal the `expected` one.
    let sequenceEqual actual expected message =
      match Utils.firstDiff actual expected with
      | _,None,None -> ()
      | i,Some a, Some e ->
        failwithf "%s. Sequence does not match at position %i. Expected item: %A, but got %A."
          message i e a
      | i,None,Some e ->
        failwithf "%s. Sequence actual shorter than expected, at pos %i for expected item %A."
          message i e
      | i,Some a,None ->
        failwithf "%s. Sequence actual longer than expected, at pos %i found item %A."
          message i a

    let columnsEqual (actual : FsCell seq seq) (expected : FsCell seq seq) message =     
        let f (cols : FsCell seq seq) = 
            cols
            |> Seq.map (fun r -> r |> Seq.map (fun c -> c.Value) |> Seq.reduce (fun a b -> a + b)) 
        sequenceEqual (f actual) (f expected) $"{message}. Columns do not match"

    let workSheetEqual (actual : FsWorksheet) (expected : FsWorksheet) message =
        let f (ws : FsWorksheet) = 
            ws.RescanRows()
            ws.Rows
            |> Seq.map (fun r -> r.Cells |> Seq.map (fun c -> c.Value) |> Seq.reduce (fun a b -> a + b)) 
        if actual.Name <> expected.Name then
            failwithf $"{message}. Worksheet names do not match. Expected {expected.Name} but got {actual.Name}"
        Expect.sequenceEqual (f actual) (f expected) $"{message}. Worksheet does not match"

    let isDefaultTestObject (wb: FsWorkbook) = DefaultTestObject.isDefaultTestObject wb

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