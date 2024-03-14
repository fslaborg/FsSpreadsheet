namespace FsSpreadsheet.Py

open FsSpreadsheet
open Fable.Core
open Fable.Core.JsInterop
open Fable.Core.JS

open Thoth.Json.Python




/// This does currently not correctly work if you want to use this from js
/// https://github.com/fable-compiler/Fable/issues/3498
[<AttachMembers>]
type Json =
    
    static member fromJsonString (json:string) : FsWorkbook =
        match Decode.fromString FsSpreadsheet.Json.Workbook.decode json with
        | Ok wb -> wb
        | Error e -> failwithf "Could not deserialize json Workbook: \n%s" e
        
    static member toJsonString (wb:FsWorkbook, ?spaces) : string =
        let spaces = defaultArg spaces 2
        FsSpreadsheet.Json.Workbook.encode wb
        |> Encode.toString spaces
