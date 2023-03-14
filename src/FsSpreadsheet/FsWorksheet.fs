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
                newCell.SetValueAs(cell.Value)
                |> ignore
        newSheet

    /// Returns a copy of a given FsWorksheet.
    static member copy (sheet : FsWorksheet) =
        sheet.Copy()


    // ------
    // Row(s)
    // ------

    /// <summary>Returns the FsRow at the given index. If it does not exist, it is created and appended first.</summary>
    member self.Row(rowIndex) = 
        match _rows |> List.tryFind (fun row -> row.Index = rowIndex) with
        | Some row ->
            row
        | None -> 
            let row = FsRow(rowIndex,self.CellCollection) 
            _rows <- List.append _rows [row]
            row

    /// <summary>Returns the FsRow at the given FsRangeAddress. If it does not exist, it is created and appended first.</summary>
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
                self.Row(newRange)
        )


    // --------
    // Table(s)
    // --------

    /// Returns the FsTable with the given tableName, rangeAddress, and showHeaderRow parameters. If it does not exist yet, it gets created and appended first.
    // TO DO: Ask HLW: Is this really a good name for the method?
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

    /// <summary>Returns the FsTable of the given name from an FsWorksheet if it exists. Else returns None.</summary>
    static member tryGetTableByName tableName (sheet : FsWorksheet) =
        sheet.Tables |> List.tryFind (fun t -> t.Name = tableName)

    /// <summary>Returns the FsTable of the given name from an FsWorksheet.</summary>
    static member getTableByName tableName (sheet : FsWorksheet) =
        try (sheet.Tables |> List.tryFind (fun t -> t.Name = tableName)).Value
        with _ -> failwith $"FsTable with name {tableName} is not presen in the FsWorksheet {sheet.Name}."

    // TO DO: tryGetTableByRangeAddress

    // TO DO: getTableByRangeAddress

    /// <summary>Adds an FsTable to the FsWorksheet if an FsTable with the same name is not already attached.</summary>
    // TO DO: Ask HLW: rather printfn or failwith?
    member self.AddTable(table : FsTable) =
        if self.Tables |> List.exists (fun t -> t.Name = table.Name) then
            printfn $"FsTable {table.Name} could not be appended as an FsTable with this name is already present in the FsWorksheet {self.Name}."
        else _tables <- List.append _tables [table]
        self

    /// <summary>Adds an FsTable to the FsWorksheet if an FsTable with the same name is not already attached.</summary>
    static member addTable table (sheet : FsWorksheet) =
        sheet.AddTable table

    /// <summary>Adds a list of FsTables to the FsWorksheet. All FsTables with a name already present in the FsWorksheet are not attached.</summary>
    // TO DO: Ask HLW: rather printfn or failwith?
    member self.AddTables(tables) =
        tables |> List.iter (self.AddTable >> ignore)
        self

    /// <summary>Adds a list of FsTables to an FsWorksheet. All FsTables with a name already present in the FsWorksheet are not attached.</summary>
    static member addTables tables (sheet : FsWorksheet) =
        sheet.AddTables tables


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


    /// Adds a FsCell to the FsWorksheet. !Exception if cell address already exists!
    member self.AddCell (cell:FsCell) =
        self.CellCollection.Add cell |> ignore
        self     

    /// Adds a sequence of FsCells to the FsWorksheet. !Exception if cell address already exists!
    member self.AddCells (cells:seq<FsCell>) =
        self.CellCollection.Add cells |> ignore
        self    

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
            c.SetValueAs value |> ignore
            self
        | None -> 
            self.CellCollection.Add(rowIndex, colIndex, value) |> ignore
            self

    /// Adds a value at the given row- and columnIndex of a given FsWorksheet.
    ///
    /// If an FsCell exists at the given position, it is overwritten.
    static member setValueAt (value : 'a) rowIndex colIndex (sheet : FsWorksheet) =
        sheet.SetValueAt(value, rowIndex, colIndex)

    /// Removes the value at the given row- and columnIndex from the FsWorksheet.
    member self.RemoveCellAt(rowIndex, colIndex) =
        self.CellCollection.RemoveCellAt(int32 rowIndex, int32 colIndex)
        self

    /// Removes the value at the given row- and columnIndex from an FsWorksheet.
    static member removeCellAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.RemoveCellAt(rowIndex, colIndex)

    /// <summary>Removes the value of an FsCell at given row- and columnIndex if it exists from the FsCollection.</summary>
    /// <remarks>Does nothing if the row or column of given index does not exist.</remarks>
    /// <exception cref="System.ArgumentNullException">if columnIndex is null.</exception>
    member self.TryRemoveValueAt(rowIndex, colIndex) =
        self.CellCollection.TryRemoveValueAt(rowIndex, colIndex)

    /// <summary>Removes the value of an FsCell at given row- and columnIndex if it exists from the FsCollection of a given FsWorksheet.</summary>
    /// <remarks>Does nothing if the row or column of given index does not exist.</remarks>
    /// <exception cref="System.ArgumentNullException">if columnIndex is null.</exception>
    static member tryRemoveValueAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.TryRemoveValueAt(rowIndex, colIndex)

    /// <summary>Removes the value of an FsCell at given row- and columnIndex from the FsCollection.</summary>
    /// <exception cref="System.ArgumentNullException">if rowIndex or columnIndex is null.</exception>
    /// <exception cref="System.Generic.KeyNotFoundException">if row or column at the given index does not exist.</exception>
    member self.RemoveValueAt(rowIndex, colIndex) =
        self.CellCollection.RemoveValueAt(rowIndex, colIndex)

    /// <summary>Removes the value of an FsCell at given row- and columnIndex from the FsCollection of a given FsWorksheet.</summary>
    /// <exception cref="System.ArgumentNullException">if rowIndex or columnIndex is null.</exception>
    /// <exception cref="System.Generic.KeyNotFoundException">if row or column at the given index does not exist.</exception>
    static member removeValueAt rowIndex colIndex (sheet : FsWorksheet) =
        sheet.RemoveValueAt(rowIndex, colIndex)


    // ##########################
    // Static member
    // Adds a FsCell to the FsWorksheet. !Exception if cell address already exists!
    static member addCell (cell:FsCell) (sheet :FsWorksheet) =
        sheet.AddCell cell 

    // ##########################
    // Static member
    // Adds a sequence of FsCells to the FsWorksheet. !Exception if cell address already exists!
    static member addCells (cell:seq<FsCell>) (sheet :FsWorksheet) =
        sheet.AddCells cell 
        


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