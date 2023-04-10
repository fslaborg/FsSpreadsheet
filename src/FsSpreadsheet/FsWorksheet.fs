namespace FsSpreadsheet

// Type based on the type XLWorksheet used in ClosedXml
/// <summary>
/// Creates an FsWorksheet with the given name, FsRows, FsTables, and FsCellsCollection.
/// </summary>
type FsWorksheet (name, fsRows, fsTables, fsCellsCollection) =

    let mutable _name = name

    let mutable _rows : FsRow list = fsRows

    let mutable _tables : FsTable list = fsTables

    let mutable _cells : FsCellsCollection = fsCellsCollection

    /// <summary>
    /// Creates an empty FsWorksheet.
    /// </summary>
    new () = 
        FsWorksheet("")

    /// <summary>
    /// Creates an empty FsWorksheet with the given name.
    /// </summary>
    new (name) = 
        FsWorksheet(name, [], [], FsCellsCollection())

    // TO DO: finish
    //new (name, (fsCells : seq<FsCell>)) =
    //    let cellsCollection = FsCellsCollection()
    //    fsCells 
    //    |> Seq.iter (
    //        fun c -> cellsCollection.Add(c.Address.RowNumber, c.Address.ColumnNumber, c)
    //    )
    //    let rows = 
    //    FsWorksheet(name)


    // ----------
    // PROPERTIES
    // ----------

    /// <summary>
    /// The name of the FsWorksheet.
    /// </summary>
    member this.Name
        with get() = _name
        and set(name) = _name <- name

    /// <summary>
    /// The FsCellCollection of the FsWorksheet.
    /// </summary>
    member this.CellCollection
        with get () = _cells

    /// <summary>
    /// The FsTables of the FsWorksheet.
    /// </summary>
    member this.Tables 
        with get() = _tables
        
    /// <summary>
    /// The FsRows of the FsWorksheet.
    /// </summary>
    member this.Rows
        with get () = _rows


    // -------
    // METHODS
    // -------

    /// <summary>
    /// Returns a copy of the FsWorksheet.
    /// </summary>
    member this.Copy() =
        let fcc = this.CellCollection.Copy()
        let nam = this.Name
        let rws = this.Rows |> List.map (fun r -> r.Copy())
        let tbs = this.Tables |> List.map (fun t -> t.Copy())
        FsWorksheet(nam, rws, tbs, fcc)
        //let newSheet = FsWorksheet(self.Name)
        //self.Tables 
        //|> List.iter (
        //    fun t -> 
        //        newSheet.Table(t.Name, t.RangeAddress, t.ShowHeaderRow) 
        //        |> ignore
        //)
        //for row in (self.Rows) do
        //    let (newRow : FsRow) = newSheet.Row(row.Index)
        //    for cell in row.Cells do
        //        let newCell = newRow.Cell(cell.Address,newSheet.CellCollection)
        //        newCell.SetValueAs(cell.Value)
        //        |> ignore
        //newSheet

    /// <summary>
    /// Returns a copy of a given FsWorksheet.
    /// </summary>
    static member copy (sheet : FsWorksheet) =
        sheet.Copy()


    // ------
    // Row(s)
    // ------

    /// <summary>
    /// Returns the FsRow at the given index. If it does not exist, creates and appends it first.
    /// </summary>
    member this.Row(rowIndex) = 
        match _rows |> List.tryFind (fun row -> row.Index = rowIndex) with
        | Some row ->
            row
        | None -> 
            let row = FsRow(rowIndex, this.CellCollection) 
            _rows <- List.append _rows [row]
            row

    /// <summary>
    /// Returns the FsRow at the given FsRangeAddress. If it does not exist, creates and appends first.
    /// </summary>
    member this.Row(rangeAddress : FsRangeAddress) = 
        if rangeAddress.FirstAddress.RowNumber <> rangeAddress.LastAddress.RowNumber then
            failwithf "Row may not have a range address spanning over different row indices"
        this.Row(rangeAddress.FirstAddress.RowNumber).RangeAddress <- rangeAddress

    /// <summary>
    /// Appends an FsRow to an FsWorksheet if the rowIndex is not already taken.
    /// </summary>
    static member appendRow (row : FsRow) (sheet : FsWorksheet) =
        sheet.Row(row.Index) |> ignore
        sheet

    /// <summary>
    /// Returns the FsRows of a given FsWorksheet.
    /// </summary>
    static member getRows (sheet : FsWorksheet) = 
        sheet.Rows

    /// <summary>
    /// Returns the FsRow at the given rowIndex of an FsWorksheet.
    /// </summary>
    static member getRowAt rowIndex sheet =
        FsWorksheet.getRows sheet
        // to do: getIndex
        |> Seq.find (FsRow.getIndex >> (=) rowIndex)

    /// <summary>
    /// Returns the FsRow at the given rowIndex of an FsWorksheet if it exists, else returns None.
    /// </summary>
    static member tryGetRowAt rowIndex (sheet : FsWorksheet) =
        sheet.Rows
        |> Seq.tryFind (FsRow.getIndex >> (=) rowIndex)

    /// <summary>
    /// Returns the FsRow matching or exceeding the given rowIndex if it exists, else returns None.
    /// </summary>
    static member tryGetRowAfter rowIndex (sheet : FsWorksheet) =
        sheet.Rows
        |> List.tryFind (fun r -> r.Index >= rowIndex)

    /// <summary>
    /// Inserts an FsRow into the FsWorksheet before a reference FsRow.
    /// </summary>
    member this.InsertBefore(row : FsRow, refRow : FsRow) =
        _rows
        |> List.iter (
            fun (r : FsRow) -> 
                if r.Index >= refRow.Index then
                    r.Index <- r.Index + 1
        )
        this.Row(row.Index) |> ignore
        this

    /// <summary>
    /// Inserts an FsRow into the FsWorksheet before a reference FsRow.
    /// </summary>
    static member insertBefore row (refRow : FsRow) (sheet : FsWorksheet) =
        sheet.InsertBefore(row, refRow)

    /// <summary>
    /// Returns true if the FsWorksheet contains an FsRow with the given rowIndex.
    /// </summary>
    member this.ContainsRowAt(rowIndex) =
        this.Rows
        |> List.exists (fun t -> t.Index = rowIndex)

    /// <summary>
    /// Returns true if the FsWorksheet contains an FsRow with the given rowIndex.
    /// </summary>
    static member containsRowAt rowIndex (sheet : FsWorksheet) =
        sheet.ContainsRowAt rowIndex

    /// <summary>
    /// Returns the number of FsRows contained in the FsWorksheet.
    /// </summary>
    static member countRows (sheet : FsWorksheet) =
        sheet.Rows.Length

    /// <summary>
    /// Removes the FsRow at the given rowIndex.
    /// </summary>
    member this.RemoveRowAt(rowIndex) =
        let newRows = 
            _rows 
            |> List.filter (fun r -> r.Index <> rowIndex)
        _rows <- newRows

    /// <summary>
    /// Removes the FsRow at a given rowIndex of an FsWorksheet.
    /// </summary>
    static member removeRowAt rowIndex (sheet : FsWorksheet) =
        sheet.RemoveRowAt(rowIndex)
        sheet

    /// <summary>
    /// Removes the FsRow at a given rowIndex of the FsWorksheet if the FsRow exists.
    /// </summary>
    member this.TryRemoveAt(rowIndex) =
        if this.ContainsRowAt rowIndex then
            this.RemoveRowAt rowIndex

    /// <summary>
    /// Removes the FsRow at a given rowIndex of an FsWorksheet if the FsRow exists.
    /// </summary>
    static member tryRemoveAt rowIndex sheet =
        if FsWorksheet.containsRowAt rowIndex sheet then
            sheet.RemoveRowAt rowIndex

    /// <summary>
    /// Sorts the FsRows by their rowIndex.
    /// </summary>
    member this.SortRows() = 
        _rows <- _rows |> List.sortBy (fun r -> r.Index)

    /// <summary>
    /// Applies function f to all FsRows and returns the modified FsWorksheet.
    /// </summary>
    member this.MapRowsInPlace(f : FsRow -> FsRow) =
        let indeces = this.Rows |> List.map (fun r -> r.Index)
        indeces
        |> List.map (fun i -> f (this.Row(i)))
        |> fun res -> _rows <- res
    
    /// <summary>
    /// Applies function f in a given FsWorksheet to all FsRows and returns the modified FsWorksheet.
    /// </summary>
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

    /// <summary>
    /// Returns the highest index of any FsRow.
    /// </summary>
    member this.GetMaxRowIndex() =
        try 
            this.Rows
            |> List.maxBy (fun r -> r.Index)
        with :? System.ArgumentException -> failwith "The FsWorksheet has no FsRows."

    /// <summary>
    /// Returns the highest index of any FsRow in a given FsWorksheet.
    /// </summary>
    static member getMaxRowIndex (sheet : FsWorksheet) =
        sheet.GetMaxRowIndex()

    /// <summary>
    /// Gets the string values of the FsRow at the given 1-based rowIndex.
    /// </summary>
    member this.GetRowValuesAt(rowIndex) =
        if this.ContainsRowAt rowIndex then
            this.Row(rowIndex).Cells
            |> Seq.map (fun c -> c.Value)
        else Seq.empty

    /// <summary>
    /// Gets the string values of the FsRow at the given 1-based rowIndex of a given FsWorksheet.
    /// </summary>
    static member getRowValuesAt rowIndex (sheet : FsWorksheet) =
        sheet.GetRowValuesAt rowIndex

    /// <summary>
    /// Gets the string values at the given 1-based rowIndex of the FsRow if it exists, else returns None.
    /// </summary>
    member this.TryGetRowValuesAt(rowIndex) =
        if this.ContainsRowAt rowIndex then
            Some (this.GetRowValuesAt rowIndex)
        else None

    /// <summary>
    /// Takes an FsWorksheet and gets the string values at the given 1-based rowIndex of the FsRow if it exists, else returns None.
    /// </summary>
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

    /// <summary>
    /// Checks the cell collection and recreate the whole set of rows, so that all cells are placed in a row
    /// </summary>
    member this.RescanRows() =
        let rows = _rows |> Seq.map (fun r -> r.Index,r) |> Map.ofSeq
        _cells.GetCells()
        |> Seq.groupBy (fun c -> c.RowNumber)
        |> Seq.iter (fun (rowIndex,cells) -> 
            let newRange = 
                cells
                |> Seq.sortBy (fun c -> c.ColumnNumber)
                |> fun cells ->
                    FsAddress(rowIndex,Seq.head cells |> fun c -> c.ColumnNumber),
                    FsAddress(rowIndex,Seq.last cells |> fun c -> c.ColumnNumber)
                |> FsRangeAddress
            match Map.tryFind rowIndex rows with
            | Some row -> 
                row.RangeAddress <- newRange
            | None ->
                this.Row(newRange)
        )


    // --------
    // Table(s)
    // --------

    /// <summary>
    /// Returns the FsTable with the given tableName, rangeAddress, and showHeaderRow parameters. If it does not exist yet, creates
    /// and appends it first.
    /// </summary>
    // TO DO: Ask HLW: Is this really a good name for the method?
    member this.Table(tableName, rangeAddress : FsRangeAddress, showHeaderRow : bool) = 
        match _tables |> List.tryFind (fun table -> table.Name = name) with
        | Some table ->
            table
        | None -> 
            let table = FsTable(tableName, rangeAddress, this.CellCollection, showHeaderRow)
            _tables <- List.append _tables [table]
            table

    /// <summary>
    /// Returns the FsTable with the given tableName and rangeAddress parameters. If it does not exist yet, it creates and appends
    /// it first. ShowHeaderRow is true by default.
    /// </summary>
    member this.Table(tableName, rangeAddress) = 
        this.Table(tableName, rangeAddress, true)

    /// <summary>
    /// Returns the FsTable of the given name from an FsWorksheet if it exists. Else returns None.
    /// </summary>
    static member tryGetTableByName tableName (sheet : FsWorksheet) =
        sheet.Tables |> List.tryFind (fun t -> t.Name = tableName)

    /// <summary>
    /// Returns the FsTable of the given name from an FsWorksheet.
    /// </summary>
    static member getTableByName tableName (sheet : FsWorksheet) =
        try (sheet.Tables |> List.tryFind (fun t -> t.Name = tableName)).Value
        with _ -> failwith $"FsTable with name {tableName} is not presen in the FsWorksheet {sheet.Name}."

    // TO DO: tryGetTableByRangeAddress

    // TO DO: getTableByRangeAddress

    /// <summary>
    /// Adds an FsTable to the FsWorksheet if an FsTable with the same name is not already attached.
    /// </summary>
    // TO DO: Ask HLW: rather printfn or failwith?
    member this.AddTable(table : FsTable) =
        if this.Tables |> List.exists (fun t -> t.Name = table.Name) then
            printfn $"FsTable {table.Name} could not be appended as an FsTable with this name is already present in the FsWorksheet {this.Name}."
        else _tables <- List.append _tables [table]

    /// <summary>
    /// Adds an FsTable to the FsWorksheet if an FsTable with the same name is not already attached.
    /// </summary>
    static member addTable table (sheet : FsWorksheet) =
        sheet.AddTable table

    /// <summary>
    /// Adds a list of FsTables to the FsWorksheet. All FsTables with a name already present in the FsWorksheet are not attached.
    /// </summary>
    // TO DO: Ask HLW: rather printfn or failwith?
    member this.AddTables(tables) =
        tables |> List.iter this.AddTable

    /// <summary>
    /// Adds a list of FsTables to an FsWorksheet. All FsTables with a name already present in the FsWorksheet are not attached.
    /// </summary>
    static member addTables tables (sheet : FsWorksheet) =
        sheet.AddTables tables
        sheet

    /// <summary>
    /// Removes the given FsTable from the FsWorksheet.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the given FsTable is not present in the FsWorksheet.</exception>
    member this.RemoveTable(table : FsTable) =
        if this.Tables |> List.exists (fun t -> t.Name = table.Name) then
            _tables <- this.Tables |> List.filter (fun t -> t.Name <> table.Name)
        else failwith $"FsTable {table.Name} was not found in FsWorksheet {this.Name}"

    /// <summary>
    /// Removes the given FsTable from the given FsWorksheet.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the given FsTable is not present in the FsWorksheet.</exception>
    static member removeTable table (sheet : FsWorksheet) =
        sheet.RemoveTable table
        sheet

    /// <summary>
    /// Replaces the FsTable of the same name with the given one.
    /// </summary>
    /// <exception cref="System.ArgumentException">if an FsTable with the given name is not present in the FsWorksheet.</exception>
    member this.UpdateTable(table : FsTable) =
        if this.Tables |> List.exists (fun t -> t.Name = table.Name) then
            _tables <- this.Tables |> List.map (fun t -> if t.Name = table.Name then table else t)
        else failwith $"FsTable {table.Name} was not found in FsWorksheet {this.Name}"

    /// <summary>
    /// Replaces the FsTable of the same name with the given one in the FsWorksheet.
    /// </summary>
    /// <exception cref="System.ArgumentException">if an FsTable with the given name is not present in the FsWorksheet.</exception>
    static member updateTable table (sheet : FsWorksheet) =
        sheet.UpdateTable table
        sheet


    // -------
    // Cell(s)
    // -------

    /// <summary>
    /// Returns the FsCell at the given row- and columnIndex if the FsCell exists, else returns None.
    /// </summary>
    member this.TryGetCellAt(rowIndex, colIndex) =
        this.CellCollection.TryGetCell(rowIndex, colIndex)

    /// <summary>
    /// Returns the FsCell at the given row- and columnIndex of a given FsWorksheet if the FsCell exists, else returns None.
    /// </summary>
    static member tryGetCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.TryGetCellAt(rowIndex, colIndex)

    /// <summary>
    /// Returns the FsCell at the given row- and columnIndex.
    /// </summary>
    member this.GetCellAt(rowIndex, colIndex) =
        this.TryGetCellAt(rowIndex, colIndex).Value

    /// <summary>
    /// Returns the FsCell at the given row- and columnIndex of a given FsWorksheet.
    /// </summary>
    static member getCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.GetCellAt(rowIndex, colIndex)

    /// <summary>
    /// Adds a FsCell to the FsWorksheet. !Exception if cell address already exists!
    /// </summary>
    member this.AddCell (cell:FsCell) =
        this.CellCollection.Add cell |> ignore

    /// <summary>
    /// Adds a sequence of FsCells to the FsWorksheet. !Exception if cell address already exists!
    /// </summary>
    member this.AddCells (cells:seq<FsCell>) =
        this.CellCollection.Add cells |> ignore

    /// <summary>
    /// Adds a value at the given row- and columnIndex to the FsWorksheet.
    ///
    /// If a cell exists at the given postion, it is shoved to the right.
    /// </summary>
    member this.InsertValueAt(value : 'a, rowIndex, colIndex)=
        let cell = FsCell(value)
        this.CellCollection.Add(int32 rowIndex, int32 colIndex, cell)

    /// <summary>
    /// Adds a value at the given row- and columnIndex to a given FsWorksheet.
    ///
    /// If an FsCell exists at the given position, it is shoved to the right.
    /// </summary>
    static member insertValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet)=
        sheet.InsertValueAt(value, rowIndex, colIndex)
        sheet

    /// <summary>
    /// Adds a value at the given row- and columnIndex.
    ///
    /// If an FsCell exists at the given position, overwrites it.
    /// </summary>
    member this.SetValueAt(value : 'a, rowIndex, colIndex) =
        match this.CellCollection.TryGetCell(rowIndex, colIndex) with
        | Some c -> 
            c.SetValueAs value |> ignore
        | None -> 
            this.CellCollection.Add(rowIndex, colIndex, value) |> ignore

    /// <summary>
    /// Adds a value at the given row- and columnIndex of a given FsWorksheet.
    ///
    /// If an FsCell exists at the given position, it is overwritten.
    /// </summary>
    static member setValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet) =
        sheet.SetValueAt(value, rowIndex, colIndex)
        sheet

    /// <summary>
    /// Removes the value at the given row- and columnIndex from the FsWorksheet.
    /// </summary>
    member this.RemoveCellAt(rowIndex, colIndex) =
        this.CellCollection.RemoveCellAt(int32 rowIndex, int32 colIndex)

    /// <summary>
    /// Removes the value at the given row- and columnIndex from an FsWorksheet.
    /// </summary>
    static member removeCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.RemoveCellAt(rowIndex, colIndex)
        sheet

    /// <summary>
    /// Removes the value of an FsCell at given row- and columnIndex if it exists from the FsCollection.
    /// </summary>
    /// <remarks>Does nothing if the row or column of given index does not exist.</remarks>
    /// <exception cref="System.ArgumentNullException">if columnIndex is null.</exception>
    member this.TryRemoveValueAt(rowIndex, colIndex) =
        this.CellCollection.TryRemoveValueAt(rowIndex, colIndex)

    /// <summary>
    /// Removes the value of an FsCell at given row- and columnIndex if it exists from the FsCollection of a given FsWorksheet.
    /// </summary>
    /// <remarks>Does nothing if the row or column of given index does not exist.</remarks>
    /// <exception cref="System.ArgumentNullException">if columnIndex is null.</exception>
    static member tryRemoveValueAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.TryRemoveValueAt(rowIndex, colIndex)
        sheet

    /// <summary>
    /// Removes the value of an FsCell at given row- and columnIndex from the FsCollection.
    /// </summary>
    /// <exception cref="System.ArgumentNullException">if rowIndex or columnIndex is null.</exception>
    /// <exception cref="System.Generic.KeyNotFoundException">if row or column at the given index does not exist.</exception>
    member this.RemoveValueAt(rowIndex, colIndex) =
        this.CellCollection.RemoveValueAt(rowIndex, colIndex)

    /// <summary>
    /// Removes the value of an FsCell at given row- and columnIndex from the FsCollection of a given FsWorksheet.
    /// </summary>
    /// <exception cref="System.ArgumentNullException">if rowIndex or columnIndex is null.</exception>
    /// <exception cref="System.Generic.KeyNotFoundException">if row or column at the given index does not exist.</exception>
    static member removeValueAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.RemoveValueAt(rowIndex, colIndex)
        sheet

    /// <summary>
    /// Adds a FsCell to the FsWorksheet. !Exception if cell address already exists!
    /// </summary>
    static member addCell (cell : FsCell) (sheet : FsWorksheet) =
        sheet.AddCell cell 
        sheet

    /// <summary>
    /// Adds a sequence of FsCells to the FsWorksheet. !Exception if cell address already exists!
    /// </summary>
    static member addCells (cell : seq<FsCell>) (sheet : FsWorksheet) =
        sheet.AddCells cell 
        sheet


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