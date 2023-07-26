namespace FsSpreadsheet.Exceljs

module JsTable =

    open FsSpreadsheet
    open Fable.ExcelJs

    let fromFsTable (fscellcollection: FsCellsCollection) (fsTable: FsTable) : Table =
        let columns = 
            if fsTable.ShowHeaderRow then
                [| for headerCell in fsTable.HeadersRow().Cells(fscellcollection) do
                    yield TableColumn(headerCell.Value) |]
            else
                [||]
        let rows = 
            [| for col in fsTable.GetColumns fscellcollection do
                let cells = if fsTable.ShowHeaderRow then col.Cells |> Seq.tail else col.Cells
                yield! cells |> Seq.mapi (fun i c -> 
                    let rowValue = box c.Value
                    i, rowValue
                ) 
            |]
            |> Array.groupBy fst
            |> Array.sortBy fst
            |> Array.map (fun (_,arr) ->
                arr |> Array.map snd
            )

        Table(fsTable.Name,fsTable.RangeAddress.Range,columns,rows,headerRow = fsTable.ShowHeaderRow)