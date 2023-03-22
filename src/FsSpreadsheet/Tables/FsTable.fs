namespace FsSpreadsheet

open System.Collections.Generic

/// <summary>
/// Creates an FsTable from the given name and FsRangeAddres, with totals row shown and header row shown or not, accordingly.
/// </summary>
type FsTable (name : string, rangeAddress, showTotalsRow, showHeaderRow) = 

    inherit FsRangeBase(rangeAddress)

    let mutable _name = name

    let mutable _lastRangeAddress = rangeAddress
    let mutable _showTotalsRow : bool = showTotalsRow
    let mutable _showHeaderRow : bool = showHeaderRow

    let mutable _fieldNames : Dictionary<string,FsTableField> = Dictionary()
    let _uniqueNames : HashSet<string> = HashSet()

    /// <summary>
    /// Creates an FsTable from the given name and FsRangeAddres, with header row shown or not, accordingly.
    /// </summary>
    /// <remarks>`showTotalsRow` is false by default.</remarks>
    new (name, rangeAddress, showHeaderRow) = FsTable (name, rangeAddress, false, showHeaderRow)

    /// <summary>
    /// Creates an FsTable from the given name and FsRangeAddres.
    /// </summary>
    /// <remarks>`showTotalsRow` is false and `showHeaderRow` true, by default.</remarks>
    new (name, rangeAddress) = FsTable (name, rangeAddress, false, true)

    /// <summary>
    /// The name of the FsTable.
    /// </summary>
    member self.Name 
        with get() = _name

    /// <summary>
    /// Returns all fieldnames as `fieldname*FsTableField` dictionary.
    /// </summary>
    member self.FieldNames
        with get(cellsCollection) =
            if (_fieldNames <> null && _lastRangeAddress <> null && _lastRangeAddress.Equals(self.RangeAddress)) then 
                _fieldNames;
            else 
                _lastRangeAddress <- self.RangeAddress

                //self.RescanFieldNames(cellsCollection)
                
                _fieldNames;

    /// <summary>
    /// The FsTableFields of this FsTable.
    /// </summary>
    member self.Fields
        with get(cellsCollection) =
            let columnCount = base.ColumnCount()
            //let offset = base.RangeAddress.FirstAddress.ColumnNumber
            Seq.init columnCount (fun i -> self.GetField(i, cellsCollection))

    /// <summary>
    /// Gets or sets if the header row is shown.
    /// </summary>
    member self.ShowHeaderRow 
        with get () = _showHeaderRow
        and set(showHeaderRow) = _showHeaderRow <- showHeaderRow

    /// <summary>
    /// Returns the header row as FsRangeRow. Scans for fieldnames if `scanForNewFieldsNames` is true.
    /// </summary>
    member self.HeadersRow(scanForNewFieldsNames : bool) = 
        if (not self.ShowHeaderRow) then null;
        
        else 
            //if (scanForNewFieldsNames) then
        
            //    let tempResult = this.FieldNames;
            //    ()

            FsRange(base.RangeAddress).FirstRow();

    /// <summary>
    /// Returns the header row as FsRangeRow. Scans for new fieldnames.
    /// </summary>
    member self.HeadersRow() = 
        self.HeadersRow(true)

    /// Takes the respective FsCellsCollection for this FsTable and creates a new _fieldNames dictionary if the current one does not match.
    // TO DO: maybe HLW can specify above description a bit...
    member private self.RescanFieldNames(cellsCollection : FsCellsCollection) =
        printfn "Start RescanFieldNames"
        _fieldNames
        |> Seq.iter (fun kv -> printfn "Key: %s, index: %i, name: %s" kv.Key kv.Value.Index kv.Value.Name)
        if self.ShowHeaderRow then
            let oldFieldNames =  _fieldNames
            _fieldNames <- new Dictionary<string, FsTableField>()
            let headersRow = self.HeadersRow(false);
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
                    
                        name <- self.GetUniqueName("Column", cellPos + 1, true)
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
    member self.RescanRange() =
        let rangeAddress = 
            _fieldNames.Values
            |> Seq.map (fun v -> v.Column.RangeAddress)
            |> Seq.reduce (fun r1 r2 -> r1.Union(r2))
        base.RangeAddress <- rangeAddress

    /// <summary>
    /// Updates the FsRangeAddress of a given FsTable according to the FsTableFields associated.
    /// </summary>
    static member rescanRange (table : FsTable) =
        table.RescanRange()
        table

    /// <summary>
    /// Returns a unique name consisting of the original name and an initial offset that is raised 
    /// if the original name with that offset is already present.
    /// </summary>
    /// <param name="enforceOffset">If true, the initial offset is always applied.</param>
    // TO DO: HLW: make this description more precise. What is this method even about?
    member this.GetUniqueName(originalName : string, initialOffset : int32, enforceOffset : bool) =
        let mutable name = originalName + if enforceOffset then string initialOffset else ""
        if _uniqueNames.Contains(name) then
        
            let mutable i = initialOffset
            name <- originalName + string i
            while _uniqueNames.Contains(name) do

                i <- i + 1
                name <- originalName + string i

        _uniqueNames.Add name |> ignore
        name

    static member getUniqueNames originalName initialOffset enforceOffset (table : FsTable) =
        table.GetUniqueName(originalName, initialOffset, enforceOffset)

    /// <summary>
    /// Creates and adds FsTableFields from a sequence of field names to the FsTable.
    /// </summary>
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
    static member addFieldsFromNames (fieldNames : seq<string>) (table : FsTable) =
        table.AddFields fieldNames
        table

    /// <summary>
    /// Adds a sequence of FsTableFields to the FsTable.
    /// </summary>
    member this.AddFields(tableFields : seq<FsTableField>) =
        tableFields
        |> Seq.iter (
            fun tf -> _fieldNames.Add(tf.Name, tf)
        )

    /// <summary>
    /// Adds a sequence of FsTableFields to a given FsTable.
    /// </summary>
    static member addFields (tableFields : seq<FsTableField>) (table : FsTable) =
        table.AddFields tableFields
        table

    /// <summary>
    /// Returns the FsTableField with given name. If an FsTableField does not exist under this name in the FsTable, adds it.
    /// </summary>
    member self.Field(name : string, cellsCollection : FsCellsCollection) = 
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
                let firstAddress = FsAddress(self.RangeAddress.FirstAddress.RowNumber,self.RangeAddress.FirstAddress.ColumnNumber + offset)
                let lastAddress = FsAddress(self.RangeAddress.LastAddress.RowNumber,self.RangeAddress.FirstAddress.ColumnNumber + offset)
                FsRangeAddress(firstAddress,lastAddress)
            let column = FsRangeColumn(range)
            let newField = FsTableField(name,maxIndex + 1,column,null,null)
            if self.ShowHeaderRow then
                newField.HeaderCell(cellsCollection,true).SetValueAs name |> ignore
            _fieldNames.Add(name,newField)
            self.RescanRange()
            newField

    /// <summary>
    /// Takes a name of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of this 
    /// FsTable) and returns the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the header row has no field with the given name.</exception>
    member self.GetField(name : string, cellsCollection : FsCellsCollection) =
        let name = name.Replace("\r\n", "\n")
        try self.FieldNames(cellsCollection).Item name
        with _ -> failwith <| "The header row doesn't contain field name '" + name + "'."

    /// <summary>
    /// Takes a name of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of this 
    /// FsTable) and returns the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the header row has no field with the given name.</exception>
    static member getFieldByName (name : string) (cellsCollection : FsCellsCollection) (table : FsTable) =
        table.GetField(name, cellsCollection)

    /// <summary>
    /// Takes the index of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of
    /// this FsTable) and returns the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the FsTable has no FsTableField with the given index.</exception>
    member self.GetField(index, cellsCollection) =
        try 
            self.FieldNames(cellsCollection).Values
            |> Seq.find (fun ftf -> ftf.Index = index)
        with _ -> failwith $"FsTableField with index {index} does not exist in the FsTable."

    /// <summary>
    /// Takes a name of an FsTableField and an FsCellsCollection (belonging to the FsWorksheet of 
    /// this FsTable) and returns the index of the respective FsTableField.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the header row has no field with the given name.</exception>
    member self.GetFieldIndex(name : string, cellsCollection) =
        self.GetField(name, cellsCollection).Index

    /// <summary>
    /// Renames a fieldname of the FsTable if it exists. Else fails.
    /// </summary>
    /// <exception cref="System.ArgumentException">if the FsTableField does not exist in the FsTable.</exception>
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
    static member renameField oldName newName (table : FsTable) =
        table.RenameField(oldName, newName)
        table

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given colum index if the cell exists. Else returns None.
    /// </summary>
    member this.TryGetHeaderCellOfColumn(cellsCollection : FsCellsCollection, colIndex : int) =
        let fstRowIndex = this.RangeAddress.FirstAddress.RowNumber
        cellsCollection.GetCellsInColumn colIndex
        |> Seq.tryFind (fun c -> c.RowNumber = fstRowIndex)

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given column index in a given FsTable if the cell exists. Else
    /// returns None.
    /// </summary>
    static member tryGetHeaderCellOfColumnIndex cellsCollection (colIndex : int) (table : FsTable) =
        table.TryGetHeaderCellOfColumn(cellsCollection, colIndex)

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection if the cell exists. Else returns None.
    /// </summary>
    member this.TryGetHeaderCellOfColumn(cellsCollection : FsCellsCollection, column : FsRangeColumn) =
        this.TryGetHeaderCellOfColumn(cellsCollection, column.Index)

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection in a given FsTable if the cell exists.
    /// Else returns None.
    /// </summary>
    static member tryGetHeaderCellOfColumn cellsCollection (column : FsRangeColumn) (table : FsTable) =
        table.TryGetHeaderCellOfColumn(cellsCollection, column)

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given colum index.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    member this.GetHeaderCellOfColumn(cellsCollection, colIndex : int) =
        this.TryGetHeaderCellOfColumn(cellsCollection, colIndex).Value

    /// <summary>
    /// Returns the header cell from a given FsCellsCollection with the given colum index in a given FsTable.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    static member getHeaderCellOfColumnIndex cellsCollection (colIndex : int) (table : FsTable) =
        table.GetHeaderCellOfColumn(cellsCollection, colIndex)

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    member this.GetHeaderCellOfColumn(cellsCollection : FsCellsCollection, column : FsRangeColumn) =
        this.TryGetHeaderCellOfColumn(cellsCollection, column).Value

    /// <summary>
    /// Returns the header cell of a given FsRangeColumn from a given FsCellsCollection in a given FsTable.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    static member getHeaderCellOfColumn cellsCollection (column : FsRangeColumn) (table : FsTable) =
        table.GetHeaderCellOfColumn(cellsCollection, column)

    /// <summary>
    /// Returns the header cell of a given FsTableField from a given FsCellsCollection.
    /// </summary>
    member this.GetHeaderCellOfTableField(cellsCollection, tableField : FsTableField) =
        tableField.HeaderCell(cellsCollection, this.ShowHeaderRow)

    /// <summary>
    /// Returns the header cell of a given FsTableField from a given FsCellsCollection in a given FsTable.
    /// </summary>
    static member getHeaderCellOfTableField cellsCollection (tableField : FsTableField) (table : FsTable) =
        table.GetHeaderCellOfTableField(cellsCollection, tableField)

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection if the cell exists.
    /// Else returns None.
    /// </summary>
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
    static member tryGetHeaderCellOfTableFieldIndex cellsCollection (tableFieldIndex : int) (table : FsTable) =
        table.TryGetHeaderCellOfTableField(cellsCollection, tableFieldIndex)

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    member this.GetHeaderCellOfTableField(cellsCollection, tableFieldIndex : int) =
        this.TryGetHeaderCellOfTableField(cellsCollection, tableFieldIndex).Value

    /// <summary>
    /// Returns the header cell from an FsTableField with the given index using a given FsCellsCollection in a given FsTable.
    /// </summary>
    /// <exception cref="System.NullReferenceException">if the FsCell cannot be found.</exception>
    static member getHeaderCellOfTableFieldIndex cellsCollection (tableFieldIndex : int) (table : FsTable) =
        table.GetHeaderCellOfTableField(cellsCollection, tableFieldIndex)

    /// <summary>
    /// Returns the header cell from an FsTableField with the given name using an FsCellsCollection in the FsTable if the cell exists.
    /// Else returns None.
    /// </summary>
    member this.TryGetHeaderCellByFieldName(cellsCollection, fieldName : string) =
        match Dictionary.tryGet fieldName _fieldNames with
        | Some tf -> Some (tf.HeaderCell(cellsCollection, this.ShowHeaderRow))
        | None -> None

    /// <summary>
    /// Returns the header cell from an FsTableField with the given name using an FsCellsCollection in a given FsTable if the cell exists.
    /// Else returns None.
    /// </summary>
    static member tryGetHeaderCellByFieldName cellsCollection (fieldName : string) (table : FsTable) =
        table.TryGetHeaderCellByFieldName(cellsCollection, fieldName)

    /// <summary>
    /// Returns the data cells from a given FsCellsCollection with the given colum index.
    /// </summary>
    member this.GetDataCellsOfColumn(cellsCollection : FsCellsCollection, colIndex) =
        let fstRowIndex = this.RangeAddress.FirstAddress.RowNumber
        let lstRowIndex = this.RangeAddress.LastAddress.RowNumber
        [fstRowIndex + 1 .. lstRowIndex]
        |> Seq.choose (
            fun ri -> cellsCollection.TryGetCell(ri, colIndex)
        )

    /// <summary>
    /// Returns the data cells from a given FsCellsCollection with the given colum index in a given FsTable.
    /// </summary>
    static member getDataCellsOfColumnIndex cellsCollection (colIndex : int) (table : FsTable) =
        table.GetDataCellsOfColumn(cellsCollection, colIndex)

    // TO DO: add equivalents of the other methods regarding header cell for data cells.

    /// <summary>
    /// Creates a deep copy of this FsTable.
    /// </summary>
    member this.Copy() =
        let ra = this.RangeAddress.Copy()
        let nam = this.Name
        let shr = this.ShowHeaderRow
        FsTable(nam, ra, false, shr)

    /// <summary>
    /// Returns a deep copy of a given FsTable.
    /// </summary>
    static member copy (table : FsTable) =
        table.Copy()