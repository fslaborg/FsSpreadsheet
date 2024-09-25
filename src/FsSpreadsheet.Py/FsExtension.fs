﻿[<AutoOpenAttribute>]
module FsSpreadsheet.Py.FsSpreadsheet


#if FABLE_COMPILER_PYTHON || !FABLE_COMPILER

open FsSpreadsheet
open FsSpreadsheet.Py

// This is mainly used for fsharp based access in a fable environment. 
// If you want to use these bindings from js, you should use the ones in `Xlsx.fs`
type FsWorkbook with

    static member fromXlsxFile(path:string) : FsWorkbook =
        Xlsx.fromXlsxFile path

    static member fromXlsxStream(stream:System.IO.Stream) : FsWorkbook =
        Xlsx.fromXlsxStream stream

    static member fromXlsxBytes(bytes: byte []) : FsWorkbook =
        Xlsx.fromXlsxBytes bytes

    static member toXlsxFile(path: string) (wb:FsWorkbook) : unit =
        Xlsx.toXlsxFile path wb

    //static member toXlsxStream(stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
    //    PyWorkbook.fromFsWorkbook wb
    //    |> fun wb -> Xlsx.writeBuffer(wb,stream)

    static member toXlsxBytes(wb:FsWorkbook) : byte [] =
        Xlsx.toXlsxBytes wb

    member this.ToXlsxFile(path: string) : unit =
        FsWorkbook.toXlsxFile path this

    //member this.ToXlsxStream(stream: System.IO.Stream) : unit =
    //    FsWorkbook.toXlsxStream stream this

    member this.ToXlsxBytes() : byte [] =
        FsWorkbook.toXlsxBytes this

    static member fromRowsJsonString (json:string) : FsWorkbook =
        Json.fromRowsJsonString json

    static member toRowsJsonString (wb:FsWorkbook,?spaces,?noNumbering) : string =
        Json.toRowsJsonString(wb, ?spaces = spaces, ?noNumbering = noNumbering)

    //static member fromJsonFile (path:string) : Promise<FsWorkbook> =
    //    Json.fromJsonFile path

    //static member toJsonFile (path:string) (wb:FsWorkbook) : Promise<unit> = 
    //    Json.toJsonFile path wb

    //member this.ToJsonFile(path: string) : Promise<unit> =
    //    FsWorkbook.toJsonFile path this

    member this.ToRowsJsonString(?spaces,?noNumbering) : string =
        FsWorkbook.toRowsJsonString(this, ?spaces = spaces, ?noNumbering = noNumbering)

    static member fromColumnsJsonString (json:string) : FsWorkbook =
        Json.fromColumnsJsonString json

    static member toColumnsJsonString (wb:FsWorkbook,?spaces,?noNumbering) : string =
        Json.toColumnsJsonString(wb, ?spaces = spaces, ?noNumbering = noNumbering)

    member this.ToColumnsJsonString(?spaces,?noNumbering) : string =
        FsWorkbook.toColumnsJsonString(this, ?spaces = spaces, ?noNumbering = noNumbering)
   
#endif