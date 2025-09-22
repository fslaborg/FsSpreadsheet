﻿namespace FsSpreadsheet.Py

module PyWorkbook =

    open Fable.Core
    open FsSpreadsheet
    open Fable.Openpyxl
    open Fable.Core.PyInterop

    let fromFsWorkbook (fsWB: FsWorkbook) : Workbook =
        FsWorkbook.validateForWrite fsWB
        if fsWB.GetWorksheets().Count = 0 then 
           failwith "Workbook must contain at least one worksheet"
        let pyWB = Workbook.create()
        pyWB.remove(pyWB.active)
        fsWB.GetWorksheets()
        |> Seq.iter (fun ws -> 
            PyWorksheet.fromFsWorksheet pyWB ws |> ignore   
        )
        //if fsWB.TryGetWorksheetByName("Sheet").IsNone then
        //    pyWB
        pyWB

    let toFsWorkbook(pyWB:Workbook) : FsWorkbook =
        let fsWB = new FsWorkbook()
        pyWB.worksheets |> Array.iter (fun (ws : Worksheet) -> 
            if ws.title <> "Sheet" && ws.values.Length <> 0 then
                let w = PyWorksheet.toFsWorksheet ws
                w.RescanRows()
                fsWB.AddWorksheet(w) |> ignore
        )
        fsWB
