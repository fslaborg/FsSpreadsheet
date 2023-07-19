module TestingUtils

open FsSpreadsheet
open Expecto

module Expect = 

    let workSheetEqual (actual : FsWorksheet) (expected : FsWorksheet) message =
        let f (ws : FsWorksheet) = 
            ws.RescanRows()
            ws.Rows
            |> Seq.map (fun r -> r.Cells |> Seq.map (fun c -> c.Value) |> Seq.reduce (fun a b -> a + b)) 
        if actual.Name <> expected.Name then
            failwithf $"{message}. Worksheet names do not match. Expected {expected.Name} but got {actual.Name}"
        Expect.sequenceEqual (f actual) (f expected) $"{message}. Worksheet does not match"

    let columnsEqual (actual : FsCell seq seq) (expected : FsCell seq seq) message =     
        let f (cols : FsCell seq seq) = 
            cols
            |> Seq.map (fun r -> r |> Seq.map (fun c -> c.Value) |> Seq.reduce (fun a b -> a + b)) 
        Expect.sequenceEqual (f actual) (f expected) $"{message}. Columns do not match"