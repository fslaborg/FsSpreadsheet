namespace FsSpreadsheet.Interactive

open FsSpreadsheet
open Giraffe.ViewEngine

module Formatters =
    
    let formatWorkbook (workbook: FsWorkbook) =
        div [] [
            style [] [str ".fs-table, .fs-th, .fs-td { border: 1px solid black !important; text-align: left; border-collapse: collapse;}"]
            h3 [] [ str $"FsWorkbook with {workbook.GetWorksheets().Length} worksheets:" ]
            table [_class "fs-table"] [
                thead [] [
                    tr [] [
                        th [_class "fs-th"] []
                        th [_class "fs-th"] [ str "Worksheet" ]
                    ]
                ]
                tbody [] [
                    for (i,worksheet) in (workbook.GetWorksheets() |> Seq.indexed) do
                        tr [] [
                            td [_class "fs-td"] [ str (string (i+1)) ]
                            td [_class "fs-td"] [ str worksheet.Name ]
                        ]
                ]
            ]
        ]
        |> RenderView.AsString.htmlNode


    let formatWorksheet (worksheet: FsWorksheet) =

        // this is basically the "toSparseMatrix" function for the FsWorksheet module, might be able to refactor
        let cells =
            worksheet.CellCollection.GetCells()
            |> Seq.map (fun c -> 
                c.RowNumber - 1, c.ColumnNumber - 1, c.Value
            )
        let matrix = 
            FsSparseMatrix.init 
                "" 
                (worksheet.CellCollection.MaxRowNumber)
                (worksheet.CellCollection.MaxColumnNumber)
                cells
            |> FsSparseMatrix.toArray2D

        div [] [
            style [] [str ".fs-table, .fs-th, .fs-td { border: 1px solid black !important; text-align: left; border-collapse: collapse;}"]
            h3 [] [ str worksheet.Name ]
            table [
                _class "fs-table"
            ] [
                thead [] [
                    for i in 0 .. worksheet.CellCollection.MaxColumnNumber ->
                        th [_class "fs-th"] [ str (string (CellReference.indexToColAdress (uint32 i))) ]
                ]
                tbody [] (
                    List.init (worksheet.CellCollection.MaxRowNumber) (fun i ->
                        tr [] [
                            td [_class "fs-td"] [str (string (i+1))]
                            yield! 
                                List.init (worksheet.CellCollection.MaxColumnNumber) (fun j -> 
                                    td [_class "fs-td"] [ str matrix.[i,j] ]
                                )
                        ]
                    )
                )
            ]
        ]
        |> RenderView.AsString.htmlNode