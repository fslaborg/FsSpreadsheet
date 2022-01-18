namespace FsSpreadsheet

open System.Collections.Generic

type FsTable (name : string, rangeAddress, showTotalsRow, showHeaderRow) = 

    inherit FsRangeBase(rangeAddress)

    let mutable _name = name

    let mutable _lastRangeAddress = rangeAddress
    let mutable _showTotalsRow : bool = showTotalsRow
    let mutable _showHeaderRow : bool = showHeaderRow

    let mutable _fieldNames : Dictionary<string,FsTableField> = Dictionary()
    let _uniqueNames : HashSet<string> = HashSet()

    new (name, rangeAddress, showHeaderRow) = FsTable (name, rangeAddress, false, showHeaderRow)

    new (name, rangeAddress) = FsTable (name, rangeAddress, false, true)

    member self.Name 
        with get() = _name

    member self.FieldNames
        with get(cells) =
            if (_fieldNames <> null && _lastRangeAddress <> null && _lastRangeAddress.Equals(self.RangeAddress)) then 
                _fieldNames;
            else 
                _lastRangeAddress <- self.RangeAddress

                //self.RescanFieldNames(cells)
                
                _fieldNames;

    member self.ShowHeaderRow 
        with get () = _showHeaderRow
        and set(showHeaderRow) = _showHeaderRow <- showHeaderRow

    member self.HeadersRow(scanForNewFieldsNames : bool) = 
        if (not self.ShowHeaderRow) then null;
        
        else 
            //if (scanForNewFieldsNames) then
        
            //    let tempResult = this.FieldNames;
            //    ()

            FsRange(base.RangeAddress).FirstRow();

    member self.HeadersRow() = 
        self.HeadersRow(true)


    member private self.RescanFieldNames (cells : FsCellsCollection) =
        printfn "Start RescanFieldNames"
        _fieldNames
        |> Seq.iter (fun kv -> printfn "Key: %s, index: %i, name: %s" kv.Key kv.Value.Index kv.Value.Name)
        if self.ShowHeaderRow then
            let oldFieldNames =  _fieldNames
            _fieldNames <- new Dictionary<string, FsTableField>()
            let headersRow = self.HeadersRow(false);
            let mutable cellPos = 0
            for cell in headersRow.Cells(cells) do
                let mutable name = cell.GetString();
                match Dictionary.tryGet name oldFieldNames with
                | Some tableField ->
                    tableField.Index <- cellPos
                    _fieldNames.Add(name,tableField)
                    cellPos <- cellPos + 1
                | None -> 

                    // Be careful here. Fields names may actually be whitespace, but not empty
                    if (name = null) <> (name = "") then
                    
                        name <- self.GetUniqueName("Column", cellPos + 1, true)
                        cell.SetValue(name) |> ignore
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


    member self.RescanRange () =
        let rangeAddress = 
            _fieldNames.Values
            |> Seq.map (fun v -> v.Column.RangeAddress)
            |> Seq.reduce (fun r1 r2 -> r1.Union(r2))
        base.RangeAddress <- rangeAddress

    member this.GetUniqueName(originalName : string, initialOffset : int32, enforceOffset : bool) =
        let mutable name = originalName + if enforceOffset then string initialOffset else ""
        if _uniqueNames.Contains(name) then
        
            let mutable i = initialOffset
            name <- originalName + string i
            while _uniqueNames.Contains(name) do

                i <- i + 1
                name <- originalName + string i

        name

    //member this.AddFields(fieldNames : IEnumerable<string>) =
    
    //    _fieldNames = new Dictionary<String, IXLTableField>();

    //    Int32 cellPos = 0;
    //    foreach (var name in fieldNames)
    //    {
    //        _fieldNames.Add(name, new XLTableField(this, name) { Index = cellPos++ });
    //    }

    member self.Field(name,cells : FsCellsCollection) = 
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
                newField.HeaderCell(cells,true).SetValue name |> ignore
            _fieldNames.Add(name,newField)
            self.RescanRange()
            newField


    member this.RenameField(oldName : string, newName : string) = 
        match Dictionary.tryGet oldName _fieldNames with
        | Some field -> 
            _fieldNames.Remove(oldName) |> ignore
            _fieldNames.Add(newName, field)
        | None -> 
            raise (System.ArgumentException("The field does not exist in this table", "oldName"))

        
