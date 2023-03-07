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

    /// Returns a copy of the FsWorksheet.
    member self.Copy() =
        let newSheet = FsWorksheet(self.Name)
        self.Tables 
        |> List.iter (
            fun t -> 
                newSheet.Table(t.Name, t.RangeAddress, t.ShowHeaderRow) 
                |> ignore
        )
        for row in (self.Rows) do
            let (newRow : FsRow) = newSheet.Row(row.Index)
            for cell in row.Cells do
                let newCell = newRow.Cell(cell.Address,newSheet.CellCollection)
                newCell.SetValue(cell.Value)
                |> ignore
        newSheet

    /// Returns a copy of a given FsWorksheet.
    static member copy (sheet : FsWorksheet) =
        sheet.Copy()


    // ------
    // Row(s)
    // ------

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

    /// Inserts an FsRow into the FsWorksheet before a reference FsRow.
    member self.InsertBefore(row : FsRow, refRow : FsRow) =
        _rows
        |> List.iter (
            fun (r : FsRow) -> 
                if r.Index >= refRow.Index then
                    r.Index <- r.Index + 1
        )
        self.Row(row.Index) |> ignore
        self

    /// Inserts an FsRow into the FsWorksheet before a reference FsRow.
    static member insertBefore row (refRow : FsRow) (sheet : FsWorksheet) =
        sheet.InsertBefore(row, refRow)

    /// Returns true if the FsWorksheet contains an FsRow with the given rowIndex.
    member self.ContainsRowAt(rowIndex) =
        self.Rows
        |> List.exists (fun t -> t.Index = rowIndex)

    /// Returns true if the FsWorksheet contains an FsRow with the given rowIndex.
    static member containsRowAt rowIndex (sheet : FsWorksheet) =
        sheet.ContainsRowAt rowIndex

    /// Returns the number of FsRows contained in the FsWorksheet.
    static member countRows (sheet : FsWorksheet) =
        sheet.Rows.Length

    /// Removes the FsRow at the given rowIndex.
    member self.RemoveRowAt(rowIndex) =
        let newRows = 
            _rows 
            |> List.filter (fun r -> r.Index <> rowIndex)
        _rows <- newRows

    /// Removes the FsRow at a given rowIndex of an FsWorksheet.
    static member removeRowAt rowIndex (sheet : FsWorksheet) =
        sheet.RemoveRowAt(rowIndex)
        sheet

    /// Removes the FsRow at a given rowIndex of the FsWorksheet if the FsRow exists.
    member self.TryRemoveAt(rowIndex) =
        if self.ContainsRowAt rowIndex then
            self.RemoveRowAt rowIndex
        self

    /// Removes the FsRow at a given rowIndex of an FsWorksheet if the FsRow exists.
    static member tryRemoveAt rowIndex sheet =
        if FsWorksheet.containsRowAt rowIndex sheet then
            sheet.RemoveRowAt rowIndex

    /// Sorts the FsRows by their rowIndex.
    member self.SortRows() = 
        _rows <- _rows |> List.sortBy (fun r -> r.Index)

    /// Applies function f to all FsRows and returns the modified FsWorksheet.
    member self.MapRowsInPlace(f : FsRow -> FsRow) =
        let indeces = self.Rows |> List.map (fun r -> r.Index)
        indeces
        |> List.map (fun i -> f (self.Row(i)))
        |> fun res -> _rows <- res
        self
    
    /// Applies function f in a given FsWorksheet to all FsRows and returns the modified FsWorksheet.
    static member mapRowsInPlace (f : FsRow -> FsRow) (sheet : FsWorksheet) =
        sheet.MapRowsInPlace f

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

    /// Returns the highest index of any FsRow.
    member self.GetMaxRowIndex() =
        try 
            self.Rows
            |> List.maxBy (fun r -> r.Index)
        with :? System.ArgumentException -> failwith "The FsWorksheet has no FsRows."

    /// Returns the highest index of any FsRow in a given FsWorksheet.
    static member getMaxRowIndex (sheet : FsWorksheet) =
        sheet.GetMaxRowIndex()

    /// Gets the string values of the FsRow at the given 1-based rowIndex.
    member self.GetRowValuesAt(rowIndex) =
        if self.ContainsRowAt rowIndex then
            self.Row(rowIndex).Cells
            |> Seq.map (fun c -> c.Value)
        else Seq.empty

    /// Gets the string values of the FsRow at the given 1-based rowIndex of a given FsWorksheet.
    static member getRowValuesAt rowIndex (sheet : FsWorksheet) =
        sheet.GetRowValuesAt rowIndex

    /// Gets the string values at the given 1-based rowIndex of the FsRow if it exists, else returns None.
    member self.TryGetRowValuesAt(rowIndex) =
        if self.ContainsRowAt rowIndex then
            Some (self.GetRowValuesAt rowIndex)
        else None

    /// Takes an FsWorksheet and gets the string values at the given 1-based rowIndex of the FsRow if it exists, else returns None.
    static member tryGetRowValuesAt rowIndex (sheet : FsWorksheet) =
        sheet.TryGetRowValuesAt rowIndex

    // TO DO (later)
    ///// If an FsRow with index rowIndex exists in the FsWorksheet, moves it downwards by amount. Negative amounts will move the FsRow upwards.
    //static member moveRowVertical amount rowIndex (sheet : FsWorksheet) =
    //    match FsWorksheet.containsRowAt

    // TO DO (later)
    //static member moveRowBlockDownward rowIndex sheet =

    

    // TO DO (later)
    //static member tryGetIndexedRowValuesAt rowIndex sheet =

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


    // -------
    // Cell(s)
    // -------

    /// Returns the FsCell at the given row- and columnIndex if the FsCell exists, else returns None.
    member self.TryGetCellAt(rowIndex, colIndex) =
        self.CellCollection.TryGetCell(rowIndex, colIndex)

    /// Returns the FsCell at the given row- and columnIndex of a given FsWorksheet if the FsCell exists, else returns None.
    static member tryGetCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.TryGetCellAt(rowIndex, colIndex)

    /// Returns the FsCell at the given row- and columnIndex.
    member self.GetCellAt(rowIndex, colIndex) =
        self.TryGetCellAt(rowIndex, colIndex).Value

    /// Returns the FsCell at the given row- and columnIndex of a given FsWorksheet.
    static member getCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.GetCellAt(rowIndex, colIndex)

    /// Adds a value at the given row- and columnIndex to the FsWorksheet.
    ///
    /// If a cell exists at the given postion, it is shoved to the right.
    member self.InsertValueAt(value : 'a, rowIndex, colIndex)=
        let cell = FsCell(value)
        self.CellCollection.Add(int32 rowIndex, int32 colIndex, cell)

    /// Adds a value at the given row- and columnIndex to a given FsWorksheet.
    ///
    /// If an FsCell exists at the given position, it is shoved to the right.
    static member insertValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet)=
        sheet.InsertValueAt(value, rowIndex, colIndex)

    /// Adds a value at the given row- and columnIndex.
    ///
    /// If an FsCell exists at the given position, overwrites it.
    member self.SetValueAt(value : 'a, rowIndex, colIndex) =
        match self.CellCollection.TryGetCell(rowIndex, colIndex) with
        | Some c -> 
            c.SetValue value |> ignore
            self
        | None -> 
            self.CellCollection.Add(rowIndex, colIndex, value)
            self

    /// Adds a value at the given row- and columnIndex of a given FsWorksheet.
    ///
    /// If an FsCell exists at the given position, it is overwritten.
    static member setValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet) =
        sheet.SetValueAt(value, rowIndex, colIndex)

    /// Removes the value at the given row- and columnIndex from the FsWorksheet.
    member self.RemoveCellAt(rowIndex, colIndex) =
        self.CellCollection.Remove(int32 rowIndex, int32 colIndex)
        self

    /// Removes the value at the given row- and columnIndex from an FsWorksheet.
    static member removeCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.RemoveCellAt(rowIndex, colIndex)

    // TO DO (later)
    //static member tryRemoveValueAt 

    // TO DO (later)
    //static member tryGetCellValueAt rowIndex colIndex sheet =

    // TO DO (later)
    //static member getCellValueAt rowIndex colIndex sheet =


    // ------------
    // Worksheet(s)
    // ------------

    // TO DO (later)
    //static member toSparseValueMatrix sheet =