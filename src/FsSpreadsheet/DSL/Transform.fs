namespace FsSpreadsheet.DSL

open FsSpreadsheet

[<AutoOpen>]
module Transform = 

    let splitRowsAndColumns (els : SheetElement list) =
        let rec loop inRows inColumns current (remaining : SheetElement list) agg =
            match remaining with

            | [] when inRows ->      ("Rows",current|> List.rev) :: agg
            | [] when inColumns ->  ("Columns",current|> List.rev) :: agg

            | UnindexedColumn c :: tail when inColumns      -> loop false true (UnindexedColumn c :: current) tail agg
            | IndexedColumn (i,c) :: tail when inColumns    -> loop false true (IndexedColumn (i,c) :: current) tail agg
            | UnindexedColumn c :: tail when inRows         -> loop false true [UnindexedColumn c] tail (("Rows",current|> List.rev) :: agg)
            | IndexedColumn (i,c) :: tail when inRows       -> loop false true [IndexedColumn (i,c)] tail (("Rows",current|> List.rev) :: agg)
            | UnindexedColumn c :: tail                     -> loop false true [UnindexedColumn c] tail agg
            | IndexedColumn (i,c) :: tail                   -> loop false true [IndexedColumn (i,c)] tail agg

            | UnindexedRow r :: tail when inRows            -> loop true false (UnindexedRow r :: current) tail agg
            | IndexedRow (i,r) :: tail when inRows          -> loop true false (IndexedRow (i,r) :: current) tail agg
            | UnindexedRow r :: tail when inColumns         -> loop true false [UnindexedRow r] tail (("Columns",current|> List.rev) :: agg)
            | IndexedRow (i,r) :: tail when inColumns       -> loop true false [IndexedRow (i,r)] tail (("Columns",current|> List.rev) :: agg)
            | UnindexedRow r :: tail                        -> loop true false [UnindexedRow r] tail agg
            | IndexedRow (i,r) :: tail                      -> loop true false [IndexedRow (i,r)] tail agg
           
        loop false false [] els []
        |> List.rev

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
                |> Set.add 0 

            let getFillRowIndex () = 
                let mutable i = 1 
                while rowIndexSet.Contains i do
                    i <- i + 1
                rowIndexSet <- Set.add i rowIndexSet
                i
            
            let getNextRowIndex () =
                rowIndexSet
                |> Seq.max
                |> (+) 1

            let parseColumns (columns : SheetElement list) = 

                let baseRowIndex = getNextRowIndex()

                let mutable columnIndexSet = 
                    columns 
                    |> List.choose (fun col -> match col with | IndexedColumn(i,_) -> Some i.Index | _ -> None)
                    |> set
    
                let getNextColumnIndex () = 
                    let mutable i = 1 
                    while columnIndexSet.Contains i do
                        i <- i + 1
                    columnIndexSet <- Set.add i columnIndexSet
                    i

                columns
                |> List.iter (fun col ->
                    let colI,elements = 
                        match col with
                        | IndexedColumn(i,colElements) -> 
                            i.Index,colElements
                        | UnindexedColumn(colElements) -> 
                            getNextColumnIndex(),colElements
                    let mutable cellIndexSet = 
                        elements 
                        |> List.choose (fun el -> match el with | ColumnElement.IndexedCell(i,_) -> Some i.Index | _ -> None)
                        |> set
                    let getNextIndex () = 
                        let mutable i = 1 
                        while cellIndexSet.Contains i do
                            i <- i + 1
                        cellIndexSet <- Set.add i cellIndexSet
                        i
                   
                    elements
                    |> List.iter (fun el ->
                        
                        match el with 
                        | ColumnElement.IndexedCell(i,(datatype,value)) -> 
                            let row = sheet.Row(i.Index + baseRowIndex - 1)
                            rowIndexSet <- Set.add (i.Index) rowIndexSet
                            let cell = row.Cell(colI,sheet.CellCollection)
                            cell.DataType <- datatype
                            cell.Value <- value
                        | ColumnElement.UnindexedCell(datatype,value) -> 
                            let row = sheet.Row(getNextIndex () + baseRowIndex - 1)
                            rowIndexSet <- Set.add row.Index rowIndexSet
                            let cell = row.Cell(colI,sheet.CellCollection)
                            cell.DataType <- datatype
                            cell.Value <- value
                    )
                  
                )


            els
            |> splitRowsAndColumns
            |> List.iter (function
                | "Columns", l -> parseColumns l
                | "Rows", l ->
                    l
                    |> List.iter (function
                        | IndexedRow(i,rowElements) -> 
                            let row = sheet.Row(i.Index)
                            Workbook.parseRow sheet.CellCollection row rowElements                
                
                        | UnindexedRow(rowElements) -> 
                            let row = sheet.Row(getFillRowIndex())
                            Workbook.parseRow sheet.CellCollection row rowElements                   
                    )
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