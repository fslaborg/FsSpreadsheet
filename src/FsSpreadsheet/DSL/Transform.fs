namespace FsSpreadsheet.DSL

open FsSpreadsheet

[<AutoOpen>]
module Transform = 

    type Workbook with
     
        static member internal parseRow (cellCollection : FsCellsCollection) (row : FsRow) (els : RowElement list) =
            let mutable cellIndexSet = 
                els 
                |> List.choose (fun el -> match el with | RowElement.IndexedCell(i,_) -> Some i.Index | _ -> None)
                |> set
            let getNextIndex () = 
                let mutable i = 1 
                while cellIndexSet.Contains i do
                    i <- i + 1
                cellIndexSet <- Set.add i cellIndexSet
                i
            els
            |> List.iter (fun el ->
                match el with 
                | RowElement.IndexedCell(i,(datatype,value)) -> 
                    let cell = row.Cell(i.Index,cellCollection)
                    cell.DataType <- datatype
                    cell.Value <- value
                | RowElement.UnindexedCell(datatype,value) -> 
                    let cell = row.Cell(getNextIndex(),cellCollection)
                    cell.DataType <- datatype
                    cell.Value <- value
            )

        static member internal parseSheet (sheet : FsWorksheet) (els : SheetElement list) =
            let mutable rowIndexSet = 
                els 
                |> List.choose (fun el -> match el with | IndexedRow(i,_) -> Some i.Index | _ -> None)
                |> set
            let getNextIndex () = 
                let mutable i = 1 
                while rowIndexSet.Contains i do
                    i <- i + 1
                rowIndexSet <- Set.add i rowIndexSet
                i
            els
            |> List.iter (fun el ->
                match el with 
                | IndexedRow(i,rowElements) -> 
                    let row = sheet.Row(i.Index)
                    Workbook.parseRow sheet.CellCollection row rowElements                
                
                | UnindexedRow(rowElements) -> 
                    let row = sheet.Row(getNextIndex())
                    Workbook.parseRow sheet.CellCollection row rowElements                   
            )

        member self.Parse() =
            match self with
            | Workbook wbEls -> 
                let workbook = new FsWorkbook()
                wbEls
                |> List.iteri (fun i wbEl ->
                    match wbEl with
                    | UnnamedSheet sheetEls ->
                        let worksheet = FsWorksheet(sprintf "Sheet%i" (i+1))
                        Workbook.parseSheet worksheet sheetEls
                        workbook.AddWorksheet(worksheet) |> ignore

                    | NamedSheet (name,sheetEls) ->
                        let worksheet = FsWorksheet(name)
                        Workbook.parseSheet worksheet sheetEls
                        workbook.AddWorksheet(worksheet) |> ignore            
                )
                workbook