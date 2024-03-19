namespace FsSpreadsheet.Js

open FsSpreadsheet
open Fable.ExcelJs
open Fable.Core
open Fable.Core.JsInterop
open Fable.Core.JS

/// This does currently not correctly work if you want to use this from js
/// https://github.com/fable-compiler/Fable/issues/3498
[<AttachMembers>]
type Xlsx =
    static member fromXlsxFile (path:string) : Promise<FsWorkbook> =
        promise {
            let wb = ExcelJs.Excel.Workbook()
            do! wb.xlsx.readFile(path)
            let fswb = JsWorkbook.readToFsWorkbook wb
            return fswb
        }

    [<System.Obsolete("Use fromXlsxFile instead")>]
    static member fromFile (path:string) : Promise<FsWorkbook> =
        Xlsx.fromXlsxFile path

    static member fromXlsxStream (stream:System.IO.Stream) : Promise<FsWorkbook> =
        promise {
            let wb = ExcelJs.Excel.Workbook()
            do! wb.xlsx.read stream
            return JsWorkbook.readToFsWorkbook wb
        }

    [<System.Obsolete("Use fromXlsxStream instead")>]
    static member fromStream (stream:System.IO.Stream) : Promise<FsWorkbook> =
        Xlsx.fromXlsxStream stream

    static member fromXlsxBytes (bytes: byte []) : Promise<FsWorkbook> =
        promise {
            let wb = ExcelJs.Excel.Workbook()
            let uint8 = Fable.Core.JS.Constructors.Uint8Array.Create bytes
            do! wb.xlsx.load(uint8.buffer)
            return JsWorkbook.readToFsWorkbook wb
        }

    [<System.Obsolete("Use fromXlsxBytes instead")>]
    static member fromBytes (bytes: byte []) : Promise<FsWorkbook> =
        Xlsx.fromXlsxBytes bytes

    static member toXlsxFile (path: string) (wb:FsWorkbook) : Promise<unit> =
        let jswb = JsWorkbook.writeFromFsWorkbook wb
        jswb.xlsx.writeFile(path)

    [<System.Obsolete("Use toXlsxFile instead")>]
    static member toFile (path: string) (wb:FsWorkbook) : Promise<unit> =
        Xlsx.toXlsxFile path wb

    static member toXlsxStream (stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
        let jswb = JsWorkbook.writeFromFsWorkbook wb
        jswb.xlsx.write(stream)

    [<System.Obsolete("Use toXlsxStream instead")>]
    static member toStream (stream: System.IO.Stream) (wb:FsWorkbook) : Promise<unit> =
        Xlsx.toXlsxStream stream wb

    static member toXlsxBytes (wb:FsWorkbook) : Promise<byte []> =
        promise {
            let jswb = JsWorkbook.writeFromFsWorkbook wb
            let buffer = jswb.xlsx.writeBuffer()
            return !!buffer
        }
            
    [<System.Obsolete("Use toXlsxBytes instead")>]
    static member toBytes (wb:FsWorkbook) : Promise<byte []> =
        Xlsx.toXlsxBytes wb
            