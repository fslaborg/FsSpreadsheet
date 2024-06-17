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
    
    static member tryFromRowsJsonString (json:string) : Result<FsWorkbook, string> =
        Decode.fromString FsSpreadsheet.Json.Workbook.decodeRows json

    static member fromRowsJsonString (json:string) : FsWorkbook =
        match Json.tryFromRowsJsonString json with
        | Ok wb -> wb
        | Error e -> failwithf "Could not deserialize json Workbook: \n%s" e
        
    static member toRowsJsonString (wb:FsWorkbook, ?spaces) : string =
        let spaces = defaultArg spaces 2
        FsSpreadsheet.Json.Workbook.encodeRows wb
        |> Encode.toString spaces

    static member tryFromColumnsJsonString (json:string) : Result<FsWorkbook, string> =
        Decode.fromString FsSpreadsheet.Json.Workbook.decodeColumns json

    static member fromColumnsJsonString (json:string) : FsWorkbook =
        match Json.tryFromColumnsJsonString json with
        | Ok wb -> wb
        | Error e -> failwithf "Could not deserialize json Workbook: \n%s" e

    static member toColumnsJsonString (wb:FsWorkbook, ?spaces) : string =
        let spaces = defaultArg spaces 2
        FsSpreadsheet.Json.Workbook.encodeColumns wb
        |> Encode.toString spaces