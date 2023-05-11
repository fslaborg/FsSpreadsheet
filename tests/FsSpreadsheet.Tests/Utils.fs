[<AutoOpen>]
module Utils

module Helpers =
    
    let firstDiff s1 s2 =
        let s1 = Seq.append (Seq.map Some s1) (Seq.initInfinite (fun _ -> None))
        let s2 = Seq.append (Seq.map Some s2) (Seq.initInfinite (fun _ -> None))
        Seq.mapi2 (fun i s p -> i,s,p) s1 s2
        |> Seq.find (function |_,Some s,Some p when s=p -> false |_-> true)

#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif

module Expect =

    let myContains (seq:seq<'T>) (ele:'T) msg = 
        let contains = Seq.contains ele seq
        Expect.isTrue contains msg

    let myHasCountOf (seq: seq<'T>) (expectedCount: int) (selector: 'T -> bool) msg =
        let itemsFiltered = Seq.filter selector seq
        let actualCount = itemsFiltered |> Seq.length
        let isSameCount = expectedCount = actualCount
        Expect.isTrue isSameCount msg

    /// Expects the `actual` sequence to equal the `expected` one.
    let inline mySequenceEqual actual expected message =
      match Helpers.firstDiff actual expected with
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