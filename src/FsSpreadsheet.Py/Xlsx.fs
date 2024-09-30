namespace FsSpreadsheet.Py

open FsSpreadsheet
open FsSpreadsheet.Py
open Fable.Core
open Fable.Openpyxl

/// This does currently not correctly work if you want to use this from js
/// https://github.com/fable-compiler/Fable/issues/3498
[<AttachMembers>]
type Xlsx = 
    
    static member fromXlsxFile(path:string) : FsWorkbook =
        Xlsx.readFile path
        |> PyWorkbook.toFsWorkbook

    static member fromXlsxStream(stream:System.IO.Stream) : FsWorkbook =
        Xlsx.load stream
        |> PyWorkbook.toFsWorkbook

    static member fromXlsxBytes(bytes: byte []) : FsWorkbook =
        Xlsx.read bytes
        |> PyWorkbook.toFsWorkbook

    static member toXlsxFile(path: string) (wb:FsWorkbook) : unit =
        PyWorkbook.fromFsWorkbook wb
        |> fun wb -> Xlsx.writeFile(wb,path)

    //static member toXlsxStream(stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
    //    PyWorkbook.fromFsWorkbook wb
    //    |> fun wb -> Xlsx.writeBuffer(wb,stream)

    static member toXlsxBytes(wb:FsWorkbook) : byte [] =
        PyWorkbook.fromFsWorkbook wb
        |> fun wb -> Xlsx.write(wb)