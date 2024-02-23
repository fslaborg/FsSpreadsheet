namespace FsSpreadsheet.ExcelPy

module PyCell =

    open Fable.Core
    open Fable.Core.JsInterop
    open FsSpreadsheet
    open Fable.Openpyxl

    let fromFsCell (fsCell: FsCell) = 
        match fsCell.DataType with
        | Boolean   -> 
            fsCell.ValueAsBool() |> box |> Some
        | Number    -> 
            fsCell.ValueAsFloat() |> box |> Some
        | Date      -> 
            /// Here it will actually show the correct DateTime. But when writing, exceljs will apply local offset. 
            //let dt = fsCell.ValueAsDateTime() |> System.DateTimeOffset 
            ///// Therefore we add offset and it should work.
            //let dt = dt.ToUniversalTime() + dt.Offset |> box |> Some
            //dt
            fsCell.ValueAsDateTime().ToUniversalTime() |> box |> Some
        | String    -> 
            fsCell.Value |> Some
        | anyElse ->
            let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse         
            printfn "%s" msg
            
            fsCell.Value |> box |> Some 

    /// <summary>
    /// `worksheetName`, `rowIndex` and `columnIndex` are only used for debugging.
    /// </summary>
    /// <param name="worksheetName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="columnIndex"></param>
    /// <param name="jsCell"></param>
    let toFsCell worksheetName rowIndex columnIndex (pyCell: Cell) =
        //printfn "toFsCell worksheetName: %s, rowIndex: %i, columnIndex: %i, %A" worksheetName rowIndex columnIndex (pyCell.value, pyCell.cellType)
        let fsadress = FsAddress(rowIndex,columnIndex)
        let dt,v = 
            let dt,v = DataType.InferCellValue pyCell.value
            if v = "=TRUE()" || v = "=True()" then
                Boolean,box true
            elif v = "=FALSE()" || v = "=False()" then
                Boolean,box false
            else
                dt,v
        FsCell(v,dt,address = fsadress)
        


