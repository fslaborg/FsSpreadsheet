[<AutoOpenAttribute>]
module FsSpreadsheet.ExcelPy.FsSpreadsheet

open FsSpreadsheet
open FsSpreadsheet.ExcelPy
open Fable.Core
open Fable.Openpyxl

// This is mainly used for fsharp based access in a fable environment. 
// If you want to use these bindings from js, you should use the ones in `Xlsx.fs`
type FsWorkbook with

    static member fromXlsxFile(path:string) : FsWorkbook =
        Xlsx.readFile path
        |> PyWorkbook.toFsWorkbook

    static member fromXlsxStream(stream:System.IO.Stream) : FsWorkbook =
        Xlsx.load stream
        |> PyWorkbook.toFsWorkbook

    static member fromBytes(bytes: byte []) : FsWorkbook =
        Xlsx.read bytes
        |> PyWorkbook.toFsWorkbook

    static member toFile(path: string) (wb:FsWorkbook) : unit =
        PyWorkbook.fromFsWorkbook wb
        |> fun wb -> Xlsx.writeFile(wb,path)

    //static member toStream(stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
    //    PyWorkbook.fromFsWorkbook wb
    //    |> fun wb -> Xlsx.writeBuffer(wb,stream)

    static member toBytes(wb:FsWorkbook) : byte [] =
        PyWorkbook.fromFsWorkbook wb
        |> fun wb -> Xlsx.write(wb)

    member this.ToFile(path: string) : unit =
        FsWorkbook.toFile path this

    //member this.ToStream(stream: System.IO.Stream) : unit =
    //    FsWorkbook.toStream stream this

    member this.ToBytes() : byte [] =
        FsWorkbook.toBytes this