namespace FsSpreadsheet.Exceljs


module JsWorkbook =
    open FsSpreadsheet
    open Fable.ExcelJs
    open Fable.Core
    
    [<Emit("console.log($0)")>]
    let private log (obj:obj) = jsNative

    let toFsWorkbook (jswb: Workbook) =
        let fswb = new FsWorkbook()
        for jsws in jswb.worksheets do
            JsWorksheet.addJsWorksheet fswb jsws
        fswb

    let fromFsWorkbook (fswb: FsWorkbook) =
        let jswb = ExcelJs.Excel.Workbook()
        log "Hit"
        for fsws in fswb.GetWorksheets() do
            JsWorksheet.addFsWorksheet jswb fsws 
        jswb
