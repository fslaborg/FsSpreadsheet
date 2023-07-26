[<AutoOpenAttribute>]
module FsSpreadsheet.Exceljs.FsSpreadsheet

open FsSpreadsheet
open FsSpreadsheet.Exceljs
open Fable.Core
open Fable.Core.JS
open Fable.Core.JsInterop

// This is mainly used for fsharp based access in a fable environment. 
// If you want to use these bindings from js, you should use the ones in `Xlsx.fs`
type FsWorkbook with

    static member fromXlsxFile(path:string) : Async<FsWorkbook> =
        Xlsx.fromXlsxFile(path)

    static member fromXlsxStream(stream:System.IO.Stream) : Async<FsWorkbook> =
        Xlsx.fromXlsxStream stream

    static member fromBytes(bytes: byte []) : Async<FsWorkbook> =
        Xlsx.fromBytes bytes

    static member toFile(path: string) (wb:FsWorkbook) : Async<unit> =
        Xlsx.toFile path wb

    static member toStream(stream: System.IO.Stream) (wb:FsWorkbook) : Async<unit> =
        Xlsx.toStream stream wb

    static member toBytes(wb:FsWorkbook) : Async<byte []> =
        Xlsx.toBytes wb

    member this.ToFile(path: string) : Async<unit> =
        FsWorkbook.toFile path this

    member this.ToStream(stream: System.IO.Stream) : Async<unit> =
        FsWorkbook.toStream stream this

    member this.ToBytes() : Async<byte []> =
        FsWorkbook.toBytes this