namespace FsSpreadsheet.Exceljs

module JsTable =

    open Fable.Core
    open FsSpreadsheet
    open Fable.ExcelJs
    open Fable.Core.JsInterop

    [<Emit("console.log($0)")>]
    let private log (obj:obj) = jsNative

    let fromFsTable (fscellcollection: FsCellsCollection) (fsTable: FsTable) : Table =
        //fsTable.ShowHeaderRow <- true //hotfix for issue #69
        let fsColumns = fsTable.GetColumns fscellcollection
        let columns = 
            if fsTable.ShowHeaderRow then
                [| for headerCell in fsTable.HeadersRow().Cells(fscellcollection) do
                    yield TableColumn(headerCell.Value) |]
            else
                [|
                    for i in 1 .. Seq.length fsColumns do yield TableColumn(string i)
                |] 
        let rows = 
            [| for col in fsColumns do
                let cells = 
                    if fsTable.ShowHeaderRow then col.Cells |> Seq.tail else col.Cells 
                yield! cells |> Seq.mapi (fun i c -> 
                    let rowValue = 
                        match c.DataType with
                        | Boolean   -> c.ValueAsBool() |> box
                        | Number    -> c.ValueAsFloat() |> box
                        | Date      -> c.ValueAsDateTime() |> box
                        | String    -> c.Value |> box
                        | anyElse ->
                            let msg = sprintf "ValueType '%A' is not fully implemented in FsSpreadsheet and is handled as string input." anyElse
                            #if FABLE_COMPILER_JAVASCRIPT
                            log msg
                            #else
                            printfn "%s" msg
                            #endif
                            c.Value |> box
                    i+1, rowValue
                ) 
            |]
            |> Array.groupBy fst
            |> Array.sortBy fst
            |> Array.map (fun (_,arr) ->
                arr |> Array.map snd
            )
        log ("Name: ",fsTable.Name)
        log ("Range: ",fsTable.RangeAddress.Range)
        log ("columns: ",columns.Length)
        log ("columns: ",columns)
        log ("rows: ",rows.Length)
        log ("headerRow: ",fsTable.ShowHeaderRow)
        Table(fsTable.Name,fsTable.RangeAddress.Range,columns,rows,headerRow = fsTable.ShowHeaderRow)