namespace FsSpreadsheet.ExcelPy

module PyWorkbook =

    open Fable.Core
    open FsSpreadsheet
    open Fable.Openpyxl
    open Fable.Core.PyInterop

    let fromFsWorkbook (fsWB: FsWorkbook) : Workbook =
        let pyWB = Workbook.create()
        fsWB.GetWorksheets()
        |> Seq.iter (fun ws -> 
            PyWorksheet.fromFsWorksheet pyWB ws |> ignore   
        )
        pyWB

    let toFsWorkbook(pyWB:Workbook) : FsWorkbook =
        let fsWB = new FsWorkbook()
        pyWB?worksheets |> Array.iter (fun (ws : Worksheet) -> 
            if ws.title <> "Sheet" && ws.values.Length <> 0 then
                let w = PyWorksheet.toFsWorksheet ws
                fsWB.AddWorksheet(w) |> ignore
        )
        fsWB
