namespace FsSpreadsheet

// Type based on the type XLWorksheet used in ClosedXml
/// Creates an FsWorksheet with the given name, FsRows, FsTables, and FsCellsCollection.
type FsWorksheet (name, fsRows, fsTables, fsCellsCollection) =

    let mutable _name = name

    let mutable _rows : FsRow list = fsRows

    let mutable _tables : FsTable list = fsTables

    let mutable _cells : FsCellsCollection = fsCellsCollection

    /// Creates an empty FsWorksheet.
    new () = 
        FsWorksheet("")

    /// Creates an empty FsWorksheet with the given name.
    new (name) = 
        FsWorksheet(name, [], [], FsCellsCollection())


    // ----------
    // PROPERTIES
    // ----------

    /// The name of the FsWorksheet.
    member self.Name
        with get() = _name
        and set(name) = _name <- name

    /// The FsCellCollection of the FsWorksheet.
    member self.CellCollection
        with get () = _cells

    /// The FsTables of the FsWorksheet.
    member self.Tables 
        with get() = _tables
        
    /// The FsRows of the FsWorksheet.
    member self.Rows
        with get () = _rows


    // -------
    // METHODS
    // -------

    /// Returns the FsRow at the given index. If it does not exist, it is created and appended first.
    member self.Row(rowIndex) = 
        match _rows |> List.tryFind (fun row -> row.Index = rowIndex) with
        | Some row ->
            row
        | None -> 
            let row = FsRow(rowIndex,self.CellCollection) 
            _rows <- List.append _rows [row]
            row

    /// Returns the FsRow at the given FsRangeAddress. If it does not exist, it is created and appended first.
    member self.Row(rangeAddress : FsRangeAddress) = 
        if rangeAddress.FirstAddress.RowNumber <> rangeAddress.LastAddress.RowNumber then
            failwithf "Row may not have a range address spanning over different row indices"
        self.Row(rangeAddress.FirstAddress.RowNumber).RangeAddress <- rangeAddress

    /// Removes the FsRow at the given rowIndex.
    member self.RemoveRowAt(rowIndex) =
        let newRows = 
            _rows 
            |> List.filter (fun r -> r.Index <> rowIndex)
        _rows <- newRows

    /// Sorts the FsRows by their rowIndex.
    member self.SortRows() = 
        _rows <- _rows |> List.sortBy (fun r -> r.Index)

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

    /// Returns the FsTable with the given tableName, rangeAddress, and showHeaderRow parameters. If it does not exist yet, it gets created and appended first.
    member self.Table(tableName,rangeAddress,showHeaderRow) = 
        match _tables |> List.tryFind (fun table -> table.Name = name) with
        | Some table ->
            table
        | None -> 
            let table = FsTable(tableName,rangeAddress,showHeaderRow)
            _tables <- List.append _tables [table]
            table
    
    /// Returns the FsTable with the given tableName and rangeAddress parameters. If it does not exist yet, it gets created first. ShowHeaderRow is true by default.
    member self.Table(tableName,rangeAddress) = 
        self.Table(tableName,rangeAddress,true)


    // --------------
    // STATIC METHODS
    // --------------

    /// Returns a copy of a given FsWorksheet.
    static member copy (sheet : FsWorksheet) =
        let newSheet = FsWorksheet(sheet.Name)
        sheet.Tables 
        |> List.iter (
            fun t -> 
                newSheet.Table(t.Name, t.RangeAddress, t.ShowHeaderRow) 
                |> ignore
        )
        for row in (sheet.Rows) do
            let newRow = newSheet.Row(row.Index)
            for cell in row.Cells do
                let newCell = newRow.Cell(cell.Address,newSheet.CellCollection)
                newCell.SetValue(cell.Value)
                |> ignore
        newSheet


    // ------
    // Row(s)
    // ------

    /// Inserts an FsRow into the FsWorksheet before a reference FsRow.
    static member insertBefore row (refRow : FsRow) (sheet : FsWorksheet) =
        FsWorksheet.getRows sheet
        |> List.iter (
            fun (r : FsRow) -> 
                if r.Index >= refRow.Index then
                    r.Index <- r.Index + 1
        )
        FsWorksheet.appendRow row sheet

    /// Appends an FsRow to an FsWorksheet if the rowIndex is not already taken.
    static member appendRow (row : FsRow) (sheet : FsWorksheet) =
        sheet.Row(row.Index) |> ignore
        sheet

    /// Returns the FsRows of a given FsWorksheet.
    static member getRows (sheet : FsWorksheet) = 
        sheet.Rows

    /// Returns the FsRow at the given rowIndex of an FsWorksheet.
    static member getRowAt rowIndex sheet =
        FsWorksheet.getRows sheet
        // to do: getIndex
        |> Seq.find (FsRow.getIndex >> (=) rowIndex)

    /// Returns the FsRow at the given rowIndex of an FsWorksheet if it exists, else returns None.
    static member tryGetRowAt rowIndex (sheet : FsWorksheet) =
        sheet.Rows
        |> Seq.tryFind (FsRow.getIndex >> (=) rowIndex)

    /// Returns the FsRow matching or exceeding the given rowIndex if it exists, else returns None.
    static member tryGetRowAfter rowIndex (sheet : FsWorksheet) =
        sheet.Rows
        |> List.tryFind (fun r -> r.Index >= rowIndex)

    /// Returns true if the FsWorksheet contains an FsRow with the given rowIndex.
    static member containsRowAt rowIndex (sheet : FsWorksheet) =
        sheet.Rows
        |> List.exists (fun t -> t.Index = rowIndex)

    // QUESTION: Does that make any sense at all? After all, CellValues are changed inPlace anyway
    ///// Builds a new FsWorksheet whose FsRows are the result of of applying the given function f to each FsRow of the given FsWorksheet.
    //static member mapRows (f : FsRow -> FsRow) (sheet : FsWorksheet) =
    //    let newRows = List.map f sheet.Rows
    //    let newWs = FsWorksheet(sheet.Name)
    //    sheet.Tables 
    //    |> List.iter (
    //        fun t -> 
    //            newWs.Table(t.Name, t.RangeAddress, t.ShowHeaderRow) 
    //            |> ignore
    //    )
    //    newWs.Tables <- sheet.Tables
    
    /// Applies function f in a given FsWorksheet to all FsRows and returns the modified FsWorksheet.
    static member mapRowsInPlace (f : FsRow -> FsRow) (sheet : FsWorksheet) =
        let indeces = sheet.Rows |> List.map (fun r -> r.Index)
        for i in indeces do
            sheet.Row(i) <- f (sheet.Row(i))
        sheet

    /// Returns the number of FsRows contained in the FsWorksheet.
    static member countRows (sheet : FsWorksheet) =
        sheet.Rows.Length

    /// Removes the FsRow at a given rowIndex of an FsWorksheet.
    static member removeRowAt rowIndex (sheet : FsWorksheet) =
        sheet.RemoveRowAt(rowIndex)
        sheet

    /// Removes the FsRow at a given rowIndex of an FsWorksheet if the FsRow exists.
    static member tryRemoveAt rowIndex sheet =
        if FsWorksheet.containsRowAt rowIndex sheet then
            sheet.RemoveRowAt rowIndex

    // TO DO (later)
    /// If an FsRow with index rowIndex exists in the FsWorksheet, moves it downwards by amount. Negative amounts will move the FsRow upwards.
    static member moveRowVertical amount rowIndex (sheet : FsWorksheet) =
        match FsWorksheet.containsRowAt

    // TO DO (later)
    //static member moveRowBlockDownward rowIndex sheet =
        
    static member getMaxRowIndex (sheet : FsWorksheet) =
        try 
            sheet.Rows
            |> List.maxBy (fun r -> r.Index)
        with :? System.ArgumentException -> failwith "Given FsWorksheet has no FsRows."

    /// Gets the string values of the FsRow at the given 1-based rowIndex.
    static member getRowValuesAt rowIndex sheet =
        (FsWorksheet.getRowAt rowIndex sheet).Cells
        |> Seq.map (fun c -> c.Value)

    /// Gets the string values of the FsRow at the given 1-based rowIndex if it exists, else returns None.
    static member tryGetRowValuesAt rowIndex sheet =
        (FsWorksheet.tryGetRowAt rowIndex sheet)
        |> Option.map (
            fun r -> r.Cells
            >> Seq.map (fun c -> c.Value)
        )

    // TO DO (later)
    //static member tryGetIndexedRowValuesAt rowIndex sheet =

    //static member removeValueAt rowIndex column sheet =
    //    let row = self.


    // -------
    // Cell(s)
    // -------

    /// Returns the FsCell at the given row- and columnIndex of a given FsWorksheet if the FsCell exists, else returns None.
    static member tryGetCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.CellCollection.TryGetCell(rowIndex, colIndex)

    /// Returns the FsCell at the given row- and columnIndex of a given FsWorksheet.
    static member getCellAt rowIndex colIndex sheet =
        (FsWorksheet.tryGetCellAt rowIndex colIndex sheet).Value

    // TO DO (later)
    //static member tryGetCellValueAt rowIndex colIndex sheet =

    // TO DO (later)
    //static member getCellValueAt rowIndex colIndex sheet =

    
    /// Adds a value at the given row- and columnIndex to the FsWorksheet.
    ///
    /// If a cell exists at the given postion, it is shoved to the right.
    static member insertValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet)=
        //let row = FsWorksheet.getRowAt rowIndex sheet
        let cell = FsCell(value)
        sheet.CellCollection.Add(int32 rowIndex, int32 colIndex, cell)

    /// Adds a value at the given row- and columnIndex.
    ///
    /// If a cell exists at the given position, overwrites it.
    static member setValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet) =
        match sheet.CellCollection.TryGetCell(rowIndex, colIndex) with
        | Some c -> 
            c.SetValue value |> ignore
            sheet
        | None -> 
            sheet.CellCollection.Add(rowIndex, colIndex, value)
            sheet

    // TO DO (later)
    //static member tryRemoveValueAt 

    /// Removes the value at the given row- and columnIndex from the FsWorksheet.
    static member removeValueAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.CellCollection.Remove(int32 rowIndex, int32 colIndex)
        sheet


    // ------------
    // Worksheet(s)
    // ------------

    // TO DO (later)
    //static member toSparseValueMatrix sheet =