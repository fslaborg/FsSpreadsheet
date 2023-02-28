namespace FsSpreadsheet

// Type based on the type XLWorksheet used in ClosedXml
type FsWorksheet (name) =
    
    let mutable _name = name
    
    let mutable _rows : FsRow list = []

    let mutable _tables : FsTable list = []

    let mutable _cells = FsCellsCollection()

    new () = FsWorksheet("")

    member self.Name
        with get() = _name
        and set(name) = _name <- name

    member self.Row(rowIndex) = 
        match _rows |> List.tryFind (fun row -> row.Index = rowIndex) with
        | Some row ->
            row
        | None -> 
            let row = FsRow(rowIndex,self.CellCollection) 
            _rows <- List.append _rows [row]
            row

    member self.Row(rangeAddress : FsRangeAddress) = 
        if rangeAddress.FirstAddress.RowNumber <> rangeAddress.LastAddress.RowNumber then
            failwithf "Row may not have a range address spanning over different row indices"
        self.Row(rangeAddress.FirstAddress.RowNumber).RangeAddress <- rangeAddress
        
    member self.GetRows() = _rows

    member self.Table(tableName,rangeAddress,showHeaderRow) = 
        match _tables |> List.tryFind (fun table -> table.Name = name) with
        | Some table ->
            table
        | None -> 
            let table = FsTable(tableName,rangeAddress,showHeaderRow)
            _tables <- List.append _tables [table]
            table
      
    member self.Table(tableName,rangeAddress) = 
        self.Table(tableName,rangeAddress,true)


    /// Checks the cell collection and recreate the whole set of rows, so that all cells are placed in a row
    member self.RescanRows() =
        let rows = _rows |> Seq.map (fun r -> r.Index,r) |> Map.ofSeq
        _cells.GetCells()
        |> Seq.groupBy (fun c -> c.WorksheetRow)
        |> Seq.iter (fun (rowIndex,cells) -> 
            let newRange = 
                cells
                |> Seq.sortBy (fun c -> c.WorksheetColumn)
                |> fun cells ->
                    FsAddress(rowIndex,Seq.head cells |> fun c -> c.WorksheetColumn),
                    FsAddress(rowIndex,Seq.last cells |> fun c -> c.WorksheetColumn)
                |> FsRangeAddress
            match Map.tryFind rowIndex rows with
            | Some row -> 
                row.RangeAddress <- newRange
            | None ->
                self.Row(newRange)
        )

    member self.CellCollection
        with get () = _cells

    member self.GetTables() = _tables

    member self.SortRows() = _rows <- _rows |> List.sortBy (fun r -> r.Index)

    static member getRows (sheet : FsWorksheet) = 
        sheet.GetRows()

    static member getRowAt index sheet =
        FsWorksheet.getRows sheet
        // to do: getIndex
        |> Seq.find (FsRow.getIndex >> (=) index)

        

    //static member removeValueAt rowIndex column sheet =
    //    let row = self.