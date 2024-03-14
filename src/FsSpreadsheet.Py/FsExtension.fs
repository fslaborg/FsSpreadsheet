[<AutoOpenAttribute>]
module FsSpreadsheet.Py.FsSpreadsheet

open FsSpreadsheet
open FsSpreadsheet.Py
open Fable.Core

// This is mainly used for fsharp based access in a fable environment. 
// If you want to use these bindings from js, you should use the ones in `Xlsx.fs`
type FsWorkbook with

    static member fromXlsxFile(path:string) : FsWorkbook =
        Xlsx.fromXlsxFile path

    static member fromXlsxStream(stream:System.IO.Stream) : FsWorkbook =
        Xlsx.fromXlsxStream stream

    static member fromBytes(bytes: byte []) : FsWorkbook =
        Xlsx.fromBytes bytes

    static member toFile(path: string) (wb:FsWorkbook) : unit =
        Xlsx.toFile path wb

    //static member toStream(stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
    //    PyWorkbook.fromFsWorkbook wb
    //    |> fun wb -> Xlsx.writeBuffer(wb,stream)

    static member toBytes(wb:FsWorkbook) : byte [] =
        Xlsx.toBytes wb

    member this.ToFile(path: string) : unit =
        FsWorkbook.toFile path this

    //member this.ToStream(stream: System.IO.Stream) : unit =
    //    FsWorkbook.toStream stream this

    member this.ToBytes() : byte [] =
        FsWorkbook.toBytes this