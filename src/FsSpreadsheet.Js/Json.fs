namespace FsSpreadsheet.Js

#if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT || !FABLE_COMPILER


open FsSpreadsheet
open Fable.ExcelJs
open Fable.Core
open Fable.Core.JsInterop
open Fable.Core.JS

open Thoth.Json.JavaScript


/// This does currently not correctly work if you want to use this from js
/// https://github.com/fable-compiler/Fable/issues/3498
[<AttachMembers>]
type Json =
    
    static member tryFromRowsJsonString (json:string) : Result<FsWorkbook, string> =
        match Thoth.Json.JavaScript.Decode.fromString FsSpreadsheet.Json.Workbook.decodeRows json with
        | Ok wb -> Ok wb
        | Error e -> Error e

    static member fromRowsJsonString (json:string) : FsWorkbook =
        match Json.tryFromRowsJsonString json with
        | Ok wb -> wb
        | Error e -> failwithf "Could not deserialize json Workbook: \n%s" e    

    static member toRowsJsonString (wb:FsWorkbook, ?spaces, ?noNumbering) : string =
        let spaces = defaultArg spaces 2
        let noNumbering = defaultArg noNumbering false
        FsSpreadsheet.Json.Workbook.encodeRows noNumbering wb
        |> Thoth.Json.JavaScript.Encode.toString spaces

    static member tryFromColumnsJsonString (json:string) : Result<FsWorkbook, string> =
        match Thoth.Json.JavaScript.Decode.fromString FsSpreadsheet.Json.Workbook.decodeColumns json with
        | Ok wb -> Ok wb
        | Error e -> Error e

    static member fromColumnsJsonString (json:string) : FsWorkbook =
        match Json.tryFromColumnsJsonString json with
        | Ok wb -> wb
        | Error e -> failwithf "Could not deserialize json Workbook: \n%s" e

    static member toColumnsJsonString (wb:FsWorkbook, ?spaces, ?noNumbering) : string =
        let spaces = defaultArg spaces 2
        let noNumbering = defaultArg noNumbering false
        FsSpreadsheet.Json.Workbook.encodeColumns noNumbering wb
        |> Thoth.Json.JavaScript.Encode.toString spaces

    //static member fromJsonFile (path:string) : Promise<FsWorkbook> =
    //    promise {
    //        let! json = Fable.Core.JS. path
    //        return Json.fromJsonString json
    //    }

    //static member toJsonFile (path:string) (wb:FsWorkbook) : Promise<unit> =
    //    let json = Json.toJsonString wb
    //    Fable.Core.JS.writeFile path json

#endif