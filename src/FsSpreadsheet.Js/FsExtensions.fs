[<AutoOpenAttribute>]
module FsSpreadsheet.Js.FsSpreadsheet

open FsSpreadsheet
open FsSpreadsheet.Js
open Fable.Core
open Fable.Core.JS
open Fable.Core.JsInterop

// This is mainly used for fsharp based access in a fable environment. 
// If you want to use these bindings from js, you should use the ones in `Xlsx.fs`
type FsWorkbook with

    static member fromXlsxFile(path:string) : Promise<FsWorkbook> =
        Xlsx.fromXlsxFile(path)

    static member fromXlsxStream(stream:System.IO.Stream) : Promise<FsWorkbook> =
        Xlsx.fromXlsxStream stream

    static member fromXlsxBytes(bytes: byte []) : Promise<FsWorkbook> =
        Xlsx.fromXlsxBytes bytes

    static member toXlsxFile(path: string) (wb:FsWorkbook) : Promise<unit> =
        Xlsx.toXlsxFile path wb

    static member toXlsxStream(stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
        Xlsx.toXlsxStream stream wb

    static member toXlsxBytes(wb:FsWorkbook) : Promise<byte []> =
        Xlsx.toXlsxBytes wb

    member this.ToXlsxFile(path: string) : Promise<unit> =
        FsWorkbook.toXlsxFile path this

    member this.ToXlsxStream(stream: System.IO.Stream) : Promise<unit> =
        FsWorkbook.toXlsxStream stream this

    member this.ToXlsxBytes() : Promise<byte []> =
        FsWorkbook.toXlsxBytes this

     

    static member fromJsonString (json:string) : FsWorkbook =
        Json.fromJsonString json

    static member toJsonString (wb:FsWorkbook) : string =
        Json.toJsonString wb

    //static member fromJsonFile (path:string) : Promise<FsWorkbook> =
    //    Json.fromJsonFile path

    //static member toJsonFile (path:string) (wb:FsWorkbook) : Promise<unit> = 
    //    Json.toJsonFile path wb

    //member this.ToJsonFile(path: string) : Promise<unit> =
    //    FsWorkbook.toJsonFile path this

    member this.ToJsonString() : string =
        FsWorkbook.toJsonString this