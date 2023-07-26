namespace FsSpreadsheet.Exceljs

open FsSpreadsheet
open Fable.ExcelJs
open Fable.Core
open Fable.Core.JsInterop

[<AttachMembers>]
type Xlsx =
    static member fromXlsxFile (path:string) : Async<FsWorkbook> =
        async {
            let wb = ExcelJs.Excel.Workbook()
            do! wb.xlsx.readFile(path)
            return JsWorkbook.toFsWorkbook wb
        }

    static member fromXlsxStream (stream:System.IO.Stream) : Async<FsWorkbook> =
        async {
            let wb = ExcelJs.Excel.Workbook()
            do! wb.xlsx.read stream
            return JsWorkbook.toFsWorkbook wb
        }

    static member fromBytes (bytes: byte []) : Async<FsWorkbook> =
        async {
            let wb = ExcelJs.Excel.Workbook()
            let uint8 = Fable.Core.JS.Constructors.Uint8Array.Create bytes
            do! wb.xlsx.load(uint8.buffer)
            return JsWorkbook.toFsWorkbook wb
        }

    static member toFile (path: string) (wb:FsWorkbook) : Async<unit> =
        let jswb = JsWorkbook.toJsWorkbook wb
        jswb.xlsx.writeFile(path)

    static member toStream (stream: System.IO.Stream) (wb:FsWorkbook) : Async<unit> =
        let jswb = JsWorkbook.toJsWorkbook wb
        jswb.xlsx.write(stream)

    static member toBytes (wb:FsWorkbook) : Async<byte []> =
        async {
            let jswb = JsWorkbook.toJsWorkbook wb
            let buffer = jswb.xlsx.writeBuffer()
            return !!buffer
        }
            
            