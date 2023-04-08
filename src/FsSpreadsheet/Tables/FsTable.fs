namespace FsSpreadsheet

open System
open System.Collections.Generic

/// <summary>
/// Creates an FsTable from the given name, FsRangeAddress, and FsCellsCollection, with totals row shown and header row shown or
/// not, accordingly.
/// </summary>
type FsTable(name : string, rangeAddress, cellsCollection : FsCellsCollection, showTotalsRow, showHeaderRow) = 

    inherit FsRangeBase(rangeAddress)

    let mutable _name = name

    let mutable _lastRangeAddress = rangeAddress
    let mutable _showTotalsRow : bool = showTotalsRow
    let mutable _showHeaderRow : bool = showHeaderRow

    let _uniqueNames : HashSet<string> = HashSet()

    /// Adds a unique name to the set. If the name is already present, adds an increasing integer at the end.
    let addUniqueName (originalName : string) (initialOffset : int32) (enforceOffset : bool) =
        let mutable name = originalName + if enforceOffset then string initialOffset else ""
        if _uniqueNames.Contains(name) then
        
            let mutable i = initialOffset
            name <- originalName + string i
            while _uniqueNames.Contains(name) do

                i <- i + 1
                name <- originalName + string i

        _uniqueNames.Add name |> ignore
        name

    let mutable _cellsCollection =
        // input cellsColl gets cropped to match the range of the table
        let minRi = base.RangeAddress.FirstAddress.RowNumber
        let maxRi = base.RangeAddress.LastAddress.RowNumber
        let minCi = base.RangeAddress.FirstAddress.ColumnNumber
        let maxCi = base.RangeAddress.LastAddress.ColumnNumber
        let nFcc = FsCellsCollection()
        cellsCollection.GetCells()
        |> Seq.iter (
            fun fsc -> 
                match fsc.Address.RowNumber, fsc.Address.ColumnNumber with
                // cell is inside of the table range, and is header cell
                | x,y when x = minRi && y >= minCi && y <= maxCi ->
                    // header cells must have unique names (= values); name is changed according to Excel standards if duplicate
                    let newVal = addUniqueName fsc.Value 2 false
                    fsc.Value <- newVal
                    nFcc.Add fsc
                // cell is inside of the table range, but not header cell
                | x,y when x > minRi && x <= maxRi && y >= minCi && y <= maxCi
                    -> nFcc.Add fsc
                | _ -> ()
        )
        nFcc

    let mutable _fieldNames : Dictionary<string,FsTableField> = Dictionary()


    // ----------------------
    // ALTERNATE CONSTRUCTORS
    // ----------------------

    /// <summary>
    /// Creates an FsTable from the given name and FsRangeAddress, with header row shown or not, accordingly.
    /// </summary>
    /// <remarks>`showTotalsRow` is false and `cellsCollection` is empty, by default.</remarks>
    new(name, rangeAddress, showHeaderRow) = FsTable(name, rangeAddress, FsCellsCollection(), false, showHeaderRow)

    /// <summary>
    /// Creates an FsTable from the given name, FsCellsCollection, and FsRangeAddress, with header row shown or not, accordingly.
    /// </summary>
    /// <remarks>`showTotalsRow` is false by default.</remarks>
    new(name, rangeAddress, cellsCollection, showHeaderRow) = FsTable(name, rangeAddress, cellsCollection, false, showHeaderRow)

    /// <summary>
    /// Creates an FsTable from the given name and FsRangeAddress.
    /// </summary>
    /// <remarks>`showTotalsRow` is false, `cellsCollection` is empty and `showHeaderRow` true, by default.</remarks>
    new(name, rangeAddress) = FsTable(name, rangeAddress, FsCellsCollection(), false, true)

    /// <summary>
    /// Creates an FsTable from the given name, FsCellsCollection, and FsRangeAddress.
    /// </summary>
    /// <remarks>`showTotalsRow` is false, and `showHeaderRow` true, by default.</remarks>
    new(name, rangeAddress, cellsCollection) = FsTable(name, rangeAddress, cellsCollection, false, true)


    // ----------
    // PROPERTIES
    // ----------

    /// <summary>
    /// The name of the FsTable.
    /// </summary>
    member this.Name 
        with get() = _name

    /// <summary>
    /// The FsCellsCollection of this FsTable.
    /// </summary>
    member this.CellsCollection
        with get() = _cellsCollection

    /// <summary>
    /// Returns the item at the given position.
    /// </summary>
    /// <exception cref="System.ArgumentException">if an item does not exist at given position.</exception>
    member this.Item
        with get(i,j) =
            let iMin = this.GetFirstRowIndex()
            let iMax = this.GetLastRowIndex()
            let jMin = this.GetFirstColIndex()
            let jMax = this.GetLastColIndex()
            match i,j with
            // out of range case
            | x,y when x > iMax || x < iMin || y > jMax || y < jMin
                -> raise (System.IndexOutOfRangeException($"Item {i},{j} is out of range {this.RangeAddress.Range}"))
            | _ -> 
                try cellsCollection[i,j]
                // if FsCell does not exist in cellsCollection -> return empty FsCell with given address
                with :? System.ArgumentException -> 
                    let adr = FsAddress(i, j)
                    FsCell.createEmptyWithAdress adr

    /// <summary>
    /// Returns all fieldnames as `fieldname*FsTableField` dictionary.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.FieldNames
        with get(cellsCollection) =
            if (_fieldNames <> null && _lastRangeAddress <> null && _lastRangeAddress.Equals(this.RangeAddress)) then 
                _fieldNames;
            else 
                _lastRangeAddress <- this.RangeAddress

                //this.RescanFieldNames(cellsCollection)
                
                _fieldNames;

    /// <summary>
    /// The FsTableFields of this FsTable.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.Fields
        with get(cellsCollection) =
            let columnCount = base.ColumnCount()
            //let offset = base.RangeAddress.FirstAddress.ColumnNumber
            Seq.init columnCount (fun i -> this.GetField(i, cellsCollection))

    /// <summary>
    /// Gets or sets if the header row is shown.
    /// </summary>
    member this.ShowHeaderRow 
        with get () = _showHeaderRow
        and set(showHeaderRow) = _showHeaderRow <- showHeaderRow


    // -------
    // METHODS
    // -------

    /// <summary>
    /// Returns a sequence of items at the given range. If an item is not present, returns an empty FsCell with the respective address.
    /// </summary>
    /// <exception cref="System.IndexOutOfRangeException">if one or several items do not exist in given range.</exception>
    member this.GetSlice(start1, end1, start2, end2) =
        let s1 = Option.defaultValue (this.GetFirstRowIndex()) start1
        let e1 = Option.defaultValue (this.GetLastRowIndex())  end1
        let s2 = Option.defaultValue (this.GetFirstColIndex()) start2
        let e2 = Option.defaultValue (this.GetLastColIndex())  end2
        seq {
            for i = s1 to e1 do
                seq {
                    for j = s2 to e2 do this[i,j]
                }
        }

    /// <summary>
    /// Returns a sequence of items at the given range.
    /// </summary>
    /// <exception cref="System.IndexOutOfRangeException">if one or several items do not exist in given range.</exception>
    member this.GetSlice(i, start2, end2) =
        let s2 = Option.defaultValue (this.GetFirstColIndex()) start2
        let e2 = Option.defaultValue (this.GetLastColIndex())  end2
        seq {
            for j = s2 to e2 do
                this[i,j]
        }

    /// <summary>
    /// Returns a sequence of items at the given range.
    /// </summary>
    /// <exception cref="System.IndexOutOfRangeException">if one or several items do not exist in given range.</exception>
    member this.GetSlice(start1, end1, j) =
        let s1 = Option.defaultValue (this.GetFirstRowIndex()) start1
        let e1 = Option.defaultValue (this.GetLastRowIndex())  end1
        seq {
            for i = s1 to e1 do
                this[i,j]
        }

    /// <summary>
    /// Returns the header row as FsRangeRow. Scans for fieldnames if `scanForNewFieldsNames` is true.
    /// </summary>
    member this.HeadersRow(scanForNewFieldsNames : bool) = 
        if (not this.ShowHeaderRow) then null;
        
        else 
            //if (scanForNewFieldsNames) then
        
            //    let tempResult = this.FieldNames;
            //    ()

            FsRange(base.RangeAddress).FirstRow();

    /// <summary>
    /// Returns the header row as FsRangeRow. Scans for new fieldnames.
    /// </summary>
    member this.HeadersRow() = 
        this.HeadersRow(true)

    // TO DO: Ask HLW/TM how to handle this: Should the row index be determined by the current max row index (=+ 1) or derived from
    // the address of the cells themselves?
    member this.AddRow cells =
        let maxRowI = this.GetLastRowIndex()
        //let 
        //this.CellsCollection.Add
        0

    // TO DO: Ask HLW/TM how to handle this: Should the col index be determined by the current max col index (=+ 1) or derived from
    // the address of the cells themselves? How to handle the required header cell?
    member this.AddColumn cells = 0

    /// Takes the respective FsCellsCollection for this FsTable and creates a new _fieldNames dictionary if the current one does not match.
    // TO DO: maybe HLW can specify above description a bit...
    member private this.RescanFieldNames(cellsCollection : FsCellsCollection) =
        printfn "Start RescanFieldNames"
        _fieldNames
        |> Seq.iter (fun kv -> printfn "Key: %s, index: %i, name: %s" kv.Key kv.Value.Index kv.Value.Name)
        if this.ShowHeaderRow then
            let oldFieldNames =  _fieldNames
            _fieldNames <- new Dictionary<string, FsTableField>()
            let headersRow = this.HeadersRow(false);
            let mutable cellPos = 0
            for cell in headersRow.Cells(cellsCollection) do
                let mutable name = cell.Value //GetString();
                match Dictionary.tryGet name oldFieldNames with
                | Some tableField ->
                    tableField.Index <- cellPos
                    _fieldNames.Add(name,tableField)
                    cellPos <- cellPos + 1
                | None -> 

                    // Be careful here. Fields names may actually be whitespace, but not empty
                    if (name = null) <> (name = "") then    // TO DO: ask: shouldn't this be XOR?
                    
                        name <- this.GetUniqueName("Column", cellPos + 1, true)
                        cell.SetValueAs(name) |> ignore
                        cell.DataType <- DataType.String

                    if (_fieldNames.ContainsKey(name)) then
                        raise (System.ArgumentException("The header row contains more than one field name '" + name + "'."))

                    _fieldNames.Add(name, new FsTableField(name, cellPos))
                    cellPos <- cellPos + 1
        else
            
            let colCount = base.ColumnCount();
            for i = 1 to colCount (**i <= colCount**) do

                if _fieldNames.Values |> Seq.exists (fun v -> v.Index = i - 1) |> not then //.All(f => f.Index != i - 1)) then

                    let name = "Column" + string i;

                    _fieldNames.Add(name, new FsTableField(name, i - 1));

        printfn "Finished RescanFWieldNames"
        _fieldNames
        |> Seq.iter (fun kv -> printfn "Key: %s, index: %i, name: %s" kv.Key kv.Value.Index kv.Value.Name)

    /// <summary>
    /// Updates the FsRangeAddress of the FsTable according to the FsTableFields associated.
    /// </summary>
    [<Obsolete>]
    member this.RescanRange() =
        let rangeAddress = 
            _fieldNames.Values
            |> Seq.map (fun v -> v.Column.RangeAddress)
            |> Seq.reduce (fun r1 r2 -> r1.Union(r2))
        base.RangeAddress <- rangeAddress

    /// <summary>
    /// Updates the FsRangeAddress of a given FsTable according to the FsTableFields associated.
    /// </summary>
    [<Obsolete>]
    static member rescanRange (table : FsTable) =
        table.RescanRange()
        table

    /// <summary>
    /// Returns a unique name consisting of the original name and an initial offset that is raised 
    /// if the original name with that offset is already present.
    /// </summary>
    /// <param name="enforceOffset">If true, the initial offset is always applied.</param>
    // changed to private (for compatibility reasons). `addUniqueName` is now the function to use here
    member private this.GetUniqueName(originalName : string, initialOffset : int32, enforceOffset : bool) =
        let mutable name = originalName + if enforceOffset then string initialOffset else ""
        if _uniqueNames.Contains(name) then
        
            let mutable i = initialOffset
            name <- originalName + string i
            while _uniqueNames.Contains(name) do

                i <- i + 1
                name <- originalName + string i

        _uniqueNames.Add name |> ignore
        name

    //static member getUniqueName originalName initialOffset enforceOffset (table : FsTable) =
    //    table.GetUniqueName(originalName, initialOffset, enforceOffset)

    /// <summary>
    /// Creates and adds FsTableFields from a sequence of field names to the FsTable.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.AddFields(fieldNames : seq<string>) =
    //    _fieldNames = new Dictionary<String, IXLTableField>();    // let's _not_ do it this way.
        //let rangeCols = FsRangeAddress.toRangeColumns base.RangeAddress
        fieldNames
        |> Seq.iteri (
            fun i fn ->
                let tableField = FsTableField(fn, i, FsRangeColumn i)
                _fieldNames.Add(fn, tableField)
        )

    /// <summary>
    /// Creates and adds FsTableFields from a sequence of field names to a given FsTable.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member addFieldsFromNames (fieldNames : seq<string>) (table : FsTable) =
        table.AddFields fieldNames
        table

    /// <summary>
    /// Adds a sequence of FsTableFields to the FsTable.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.AddFields(tableFields : seq<FsTableField>) =
        tableFields
        |> Seq.iter (
            fun tf -> _fieldNames.Add(tf.Name, tf)
        )

    /// <summary>
    /// Adds a sequence of FsTableFields to a given FsTable.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member addFields (tableFields : seq<FsTableField>) (table : FsTable) =
        table.AddFields tableFields
        table

    /// <summary>
    /// Returns the FsTableField with given name. If an FsTableField does not exist under this name in the FsTable, adds it.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.Field(name : string, cellsCollection : FsCellsCollection) = 
        match Dictionary.tryGet name _fieldNames with
        | Some field -> 
            field
        | None -> 
            let maxIndex = 
                _fieldNames.Values 
                |> Seq.map (fun v -> v.Index) 
                |> fun s -> 
                    if Seq.length s = 0 then 0 else Seq.max s
            let range = 
                let offset = _fieldNames.Count
                let firstAddress = FsAddress(this.RangeAddress.FirstAddress.RowNumber, this.RangeAddress.FirstAddress.ColumnNumber + offset)
                let lastAddress = FsAddress(this.RangeAddress.LastAddress.RowNumber, this.RangeAddress.FirstAddress.ColumnNumber + offset)
                FsRangeAddress(firstAddress,lastAddress)
            let column = FsRangeColumn(range)
            let newField = FsTableField(name, maxIndex + 1, column, null, null)
            if this.ShowHeaderRow then
                newField.HeaderCell(cellsCollection, true).SetValueAs name |> ignore
            _fieldNames.Add(name,newField)
            this.RescanRange()
            newField

    /// <summary>
    /// Takes a name of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of this 
    /// FsTable) and returns the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the header row has no field with the given name.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.GetField(name : string, cellsCollection : FsCellsCollection) =
        let name = name.Replace("\r\n", "\n")
        try this.FieldNames(cellsCollection).Item name
        with _ -> failwith <| "The header row doesn't contain field name '" + name + "'."

    /// <summary>
    /// Takes a name of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of this 
    /// FsTable) and returns the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the header row has no field with the given name.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member getFieldByName (name : string) (cellsCollection : FsCellsCollection) (table : FsTable) =
        table.GetField(name, cellsCollection)

    /// <summary>
    /// Takes the index of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of
    /// this FsTable) and returns the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the FsTable has no FsTableField with the given index.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.GetField(index, cellsCollection) =
        try 
            this.FieldNames(cellsCollection).Values
            |> Seq.find (fun ftf -> ftf.Index = index)
        with _ -> failwith $"FsTableField with index {index} does not exist in the FsTable."

    /// <summary>
    /// Takes a name of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of 
    /// this FsTable) and returns the index of the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the header row has no field with the given name.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.GetFieldIndex(name : string, cellsCollection) =
        this.GetField(name, cellsCollection).Index

    /// <summary>
    /// Renames a fieldname of the FsTable if it exists. Else fails.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the FsTableField does not exist in the FsTable.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.RenameField(oldName : string, newName : string) = 
        match Dictionary.tryGet oldName _fieldNames with
        | Some field -> 
            _fieldNames.Remove(oldName) |> ignore
            _fieldNames.Add(newName, field)
        | None -> 
            raise (System.ArgumentException("The FsTabelField does not exist in this FsTable", "oldName"))

    /// <summary>
    /// Renames a fieldname of the FsTable if it exists. Else fails.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the FsTableField does not exist in the FsTable.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member renameField oldName newName (table : FsTable) =
        table.RenameField(oldName, newName)
        table

    /// <summary>
    /// Returns the header cell with the given colum index if the cell exists. Else returns None.
    /// </summary>
    member this.TryGetHeaderCellOfColumn(colIndex : int) =
        let fstRowIndex = this.GetFirstRowIndex()
        try this[fstRowIndex,colIndex] |> Some
        with _ -> None

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given column index in a given FsTable if the cell exists. Else
    /// returns None.
    /// </summary>
    static member tryGetHeaderCellOfColumnIndex (colIndex : int) (table : FsTable) =
        table.TryGetHeaderCellOfColumn colIndex

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection if the cell exists. Else returns None.
    /// </summary>
    member this.TryGetHeaderCellOfColumn(column : FsRangeColumn) =
        this.TryGetHeaderCellOfColumn column.Index

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection in a given FsTable if the cell exists.
    /// Else returns None.
    /// </summary>
    static member tryGetHeaderCellOfColumn (column : FsRangeColumn) (table : FsTable) =
        table.TryGetHeaderCellOfColumn column

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given colum index.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    member this.GetHeaderCellOfColumn(colIndex : int) =
        this.TryGetHeaderCellOfColumn(colIndex).Value

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given colum index in a given FsTable.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    static member getHeaderCellOfColumnIndex (colIndex : int) (table : FsTable) =
        table.GetHeaderCellOfColumn colIndex

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    member this.GetHeaderCellOfColumn(column : FsRangeColumn) =
        this.TryGetHeaderCellOfColumn(column).Value

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection in a given FsTable.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    static member getHeaderCellOfColumn (column : FsRangeColumn) (table : FsTable) =
        table.GetHeaderCellOfColumn column

    /// <summary>
    /// Returns the header cell of a given FsTableField from a given FsCellsCollection.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.GetHeaderCellOfTableField(cellsCollection, tableField : FsTableField) =
        tableField.HeaderCell(cellsCollection, this.ShowHeaderRow)

    /// <summary>
    /// Returns the header cell of a given FsTableField from a given FsCellsCollection in a given FsTable.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member getHeaderCellOfTableField cellsCollection (tableField : FsTableField) (table : FsTable) =
        table.GetHeaderCellOfTableField(cellsCollection, tableField)

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection if the cell exists.
    /// Else returns None.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.TryGetHeaderCellOfTableField(cellsCollection, tableFieldIndex : int) =
        _fieldNames.Values
        |> Seq.tryPick (
            fun tf -> 
                if tf.Index = tableFieldIndex then
                    Some (tf.HeaderCell(cellsCollection, this.ShowHeaderRow))
                else None
        )

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection if the cell exists
    /// in a given FsTable. Else returns None.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member tryGetHeaderCellOfTableFieldIndex cellsCollection (tableFieldIndex : int) (table : FsTable) =
        table.TryGetHeaderCellOfTableField(cellsCollection, tableFieldIndex)

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.GetHeaderCellOfTableField(cellsCollection, tableFieldIndex : int) =
        this.TryGetHeaderCellOfTableField(cellsCollection, tableFieldIndex).Value

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection in a given FsTable.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member getHeaderCellOfTableFieldIndex cellsCollection (tableFieldIndex : int) (table : FsTable) =
        table.GetHeaderCellOfTableField(cellsCollection, tableFieldIndex)

    /// <summary>
    /// Returns the header cell from an FsTableField with the given name using an FsCellsCollection in the FsTable if the cell exists.
    /// Else returns None.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    member this.TryGetHeaderCellByFieldName(cellsCollection, fieldName : string) =
        match Dictionary.tryGet fieldName _fieldNames with
        | Some tf -> Some (tf.HeaderCell(cellsCollection, this.ShowHeaderRow))
        | None -> None

    /// <summary>
    /// Returns the header cell from an FsTableField with the given name using an FsCellsCollection in a given FsTable if the cell exists.
    /// Else returns None.
    /// </summary>
    [<Obsolete "Use `GetHeaderCell` & `GetDataCells` instead.">]
    static member tryGetHeaderCellByFieldName cellsCollection (fieldName : string) (table : FsTable) =
        table.TryGetHeaderCellByFieldName(cellsCollection, fieldName)

    /// <summary>
    /// Returns the data cells from a given FsCellsCollection with the given colum index.
    /// </summary>
    // /// <remarks>Column index must fit the FsCellsCollection, not the FsTable!</remarks>
    member this.GetDataCellsOfColumn(colIndex) =
        let fstRowIndex = this.GetFirstRowIndex()
        let lstRowIndex = this.GetLastRowIndex()
        this[fstRowIndex + 1 .. lstRowIndex,colIndex]

    /// <summary>
    /// Returns the data cells from a given FsCellsCollection with the given colum index in a given FsTable.
    /// </summary>
    /// <remarks>Column index must fit the FsCellsCollection, not the FsTable!</remarks>
    static member getDataCellsOfColumnIndex (colIndex : int) (table : FsTable) =
        table.GetDataCellsOfColumn colIndex

    // TO DO: add equivalents of the other methods regarding header cell for data cells.

    /// <summary>
    /// Creates a deep copy of this FsTable.
    /// </summary>
    member this.Copy() =
        let ra = this.RangeAddress.Copy()
        let nam = this.Name
        let shr = this.ShowHeaderRow
        let fcc = this.CellsCollection.Copy()
        FsTable(nam, ra, fcc, false, shr)

    /// <summary>
    /// Returns a deep copy of a given FsTable.
    /// </summary>
    static member copy (table : FsTable) =
        table.Copy()

    /// <summary>
    /// Returns the index of the first row of the FsTable.
    /// </summary>
    member this.GetFirstRowIndex() =
        this.RangeAddress.FirstAddress.RowNumber

    /// <summary>
    /// Returns the index of the first row of a given FsTable.
    /// </summary>
    static member getFirstRowIndex (table : FsTable) =
        table.GetFirstRowIndex()

    /// <summary>
    /// Returns the index of the last row of the FsTable.
    /// </summary>
    member this.GetLastRowIndex() =
        this.RangeAddress.LastAddress.RowNumber

    /// <summary>
    /// Returns the index of the last row of a given FsTable.
    /// </summary>
    static member getLastRowIndex (table : FsTable) =
        table.GetFirstRowIndex()

    /// <summary>
    /// Returns the index of the first column of the FsTable.
    /// </summary>
    member this.GetFirstColIndex() =
        this.RangeAddress.FirstAddress.ColumnNumber

    /// <summary>
    /// Returns the index of the first column of a given FsTable.
    /// </summary>
    static member getFirstColIndex (table : FsTable) =
        table.GetFirstRowIndex()

    /// <summary>
    /// Returns the index of the last column of the FsTable.
    /// </summary>
    member this.GetLastColIndex() =
        this.RangeAddress.LastAddress.ColumnNumber

    /// <summary>
    /// Returns the index of the last column of a given FsTable.
    /// </summary>
    static member getLastColIndex (table : FsTable) =
        table.GetLastRowIndex()

    member this.GetHeaderCell colIndex =
        this[*,colIndex]