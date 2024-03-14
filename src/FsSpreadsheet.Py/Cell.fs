namespace FsSpreadsheet.Py

module PyCell =

    open Fable.Core
    open Fable.Core.PyInterop
    open FsSpreadsheet
    open Fable.Openpyxl

    // Currently in Fable, a created datetime object will contain a timezone. This is not allowed in python xlsx, so we need to remove it.
    // Unfortunately, the timezone object in python is read-only, so we need to create a new datetime object without timezone.
    // For this, we use the fromtimestamp method of the datetime module and convert the timestamp to a new datetime object without timezone.
    type datetime =
        abstract member decoy: unit -> unit
    type DateTimeStatic =
        [<Emit("$0.fromtimestamp(timestamp=$1)")>]
        abstract member fromTimeStamp: timestamp:float -> datetime
    [<Import("datetime", "datetime")>]
    let DateTime : DateTimeStatic = nativeOnly


    let toUniversalTimePy (dt:System.DateTime) = 
        
        dt.ToUniversalTime()?timestamp()
        |> DateTime.fromTimeStamp

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
            let dt = fsCell.ValueAsDateTime() |> toUniversalTimePy |> box
            //dt?tzinfo <- None
            dt
            |> Some
        | String    -> 
            fsCell.Value |> Some
        | anyElse ->
            let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse         
            printfn "%s" msg
            
            fsCell.Value |> box |> Some 

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
        


