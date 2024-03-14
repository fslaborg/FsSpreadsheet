namespace FsSpreadsheet.Js

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
    
    static member fromJsonString (json:string) : FsWorkbook =
        match Thoth.Json.JavaScript.Decode.fromString FsSpreadsheet.Json.Workbook.decode json with
        | Ok wb -> wb
        | Error e -> failwithf "Could not deserialize json Workbook: \n%s" e
        
    static member toJsonString (wb:FsWorkbook, ?spaces) : string =
        let spaces = defaultArg spaces 2
        FsSpreadsheet.Json.Workbook.encode wb
        |> Thoth.Json.JavaScript.Encode.toString spaces

    //static member fromJsonFile (path:string) : Promise<FsWorkbook> =
    //    promise {
    //        let! json = Fable.Core.JS. path
    //        return Json.fromJsonString json
    //    }

    //static member toJsonFile (path:string) (wb:FsWorkbook) : Promise<unit> =
    //    let json = Json.toJsonString wb
    //    Fable.Core.JS.writeFile path json