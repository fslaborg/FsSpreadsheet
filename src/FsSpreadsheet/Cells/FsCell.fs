namespace FsSpreadsheet

open System

/// <summary>
/// Possible DataTypes used in a FsCell.
/// </summary>
type DataType = 
    | String
    | Boolean
    | Number
    | Date
    | Empty

    /// <summary>
    /// Returns the proper CellValues case for the given value.
    /// </summary>
    static member InferCellValue (value : 'T) = 
        let value = box value
        match value with
        | :? char as c -> DataType.String,c.ToString()
        | :? bool as true -> DataType.Boolean, "True"
        | :? bool as false -> DataType.Boolean, "False"
        | :? byte as i -> DataType.Number,i.ToString()
        | :? sbyte as i -> DataType.Number,i.ToString()
        | :? int as i -> DataType.Number,i.ToString()
        | :? int16 as i -> DataType.Number,i.ToString()
        | :? int64 as i -> DataType.Number,i.ToString()
        | :? uint as i -> DataType.Number,i.ToString()
        | :? uint16 as i -> DataType.Number,i.ToString()
        | :? uint64 as i -> DataType.Number,i.ToString()
        | :? single as i -> DataType.Number,i.ToString()
        | :? float as i -> DataType.Number,i.ToString()
        | :? decimal as i -> DataType.Number,i.ToString()
        | :? System.DateTime as d -> DataType.Date,d.ToString()
        | :? string as s -> DataType.String,s.ToString()
        | _ ->  DataType.String,value.ToString()

// Type based on the type XLCell used in ClosedXml
/// <summary>
/// Creates an FsCell of `DataType` dataType, with value of type `string`, and `FsAddress` address.
/// </summary>
type FsCell (value : IConvertible, dataType : DataType, address : FsAddress) =
    
    // TODO: Maybe save as IConvertible
    let mutable _cellValue = string value
    let mutable _dataType = dataType
    let mutable _comment  = ""
    let mutable _hyperlink = ""
    let mutable _richText = ""
    let mutable _formulaA1 = ""
    let mutable _formulaR1C1 = ""

    let mutable _rowIndex : int = address.RowNumber
    let mutable _columnIndex : int = address.ColumnNumber


    // ------------------------
    // ALTERNATIVE CONSTRUCTORS
    // ------------------------

    new (value : IConvertible) = FsCell (string value, DataType.String, FsAddress(0,0))

    ///// Creates an empty FsCell, set at row 1, column 1 (1-based).
    //new () = FsCell ("", DataType.Empty, FsAddress(0,0))

    ///// Creates an FsCell of `DataType` `Number`, with the given value, set at row 1, column 1 (1-based).
    //new (value : int) = FsCell (string value, DataType.Number, FsAddress(0,0))

    ///// Creates an FsCell of `DataType` `Number`, with the given value, set at row 1, column 1 (1-based).
    //new (value : float) = FsCell (string value, DataType.Number, FsAddress(0,0))

    ///// Creates an empty FsCell, set at `FsAddress` address.
    //new (address : FsAddress) =
    //    FsCell ("", DataType.Empty, address)

    ///// Creates an empty FsCell, set at row rowIndex and column colIndex.
    //new (rowIndex, colIndex) =
    //    FsCell("", DataType.Empty, FsAddress(rowIndex, colIndex))


    // ----------
    // PROPERTIES
    // ----------

    /// <summary>
    /// Gets or sets the cell's value. To get or set a strongly typed value, use the GetValue&lt;T&gt; and SetValue methods.
    /// </summary>
    /// <value>
    /// The object containing the value(s) to set.
    /// </value>
    member self.Value 
        with get() = _cellValue
        and set(value) = _cellValue <- value
 
    /// <summary>
    /// Gets or sets the DataType of this FsCell's data.
    /// <para>Changing the data type will cause FsSpreadsheet to convert the current value to the new DataType. </para>
    /// <para>An exception will be thrown if the current value cannot be converted to the new DataType.</para>
    /// </summary>
    /// <value>
    /// The type of the FsCell's data.
    /// </value>
    /// <exception cref="ArgumentException"></exception>
    member self.DataType 
        with get() = _dataType
        and internal set(dataType) = _dataType <- dataType 

    /// <summary>
    /// Gets or sets the columnIndex of the FsCell.
    /// </summary>
    member self.ColumnNumber
        with get() = _columnIndex
        and set(colI) = _columnIndex <- colI
    
    /// <summary>
    /// Gets or sets the rowIndex of the FsCell.
    /// </summary>
    member self.RowNumber
        with get() = _rowIndex
        and set(rowI) = _rowIndex <- rowI

    /// <summary>
    /// Gets this FsCell's address, relative to the FsWorksheet.
    /// </summary>
    /// <value>The FsCell's address.</value>
    member self.Address 
        with get() = FsAddress(_rowIndex,_columnIndex)
        and internal set(address : FsAddress) =
            _rowIndex <- address.RowNumber
            _columnIndex <- address.ColumnNumber


    /// <summary>
    /// Create an FsCell from given rowNumber, colNumber, and value. Infers the DataType.
    /// </summary>
    static member create (rowNumber : int) (colNumber : int) value =
        let dataT, value = DataType.InferCellValue value
        FsCell(value, dataT, FsAddress(rowNumber, colNumber))

    /// <summary>
    /// Creates an empty FsCell.
    /// </summary>
    static member createEmpty ()  =
        FsCell("", DataType.Empty, FsAddress(0,0))

    /// <summary>
    /// Creates an FsCell with the given FsAdress and value. Inferes the DataType.
    /// </summary>
    static member createWithAdress (adress : FsAddress) value =
        let dataT, value = DataType.InferCellValue value
        FsCell(value, dataT, adress)

    /// <summary>
    /// Creates an empty FsCell with a given FsAddress.
    /// </summary>
    static member createEmptyWithAdress (adress : FsAddress)  =
        FsCell("", DataType.Empty, adress)

    /// <summary>
    /// Creates an FsCell with the given DataType, rowNumber, colNumber, and value.
    /// </summary>
    static member createWithDataType (dataType : DataType) (rowNumber : int) (colNumber : int) value =
        FsCell(value, dataType, FsAddress(rowNumber, colNumber))

    //how 2:
    //return (format.ToUpper()) switch
            //{
            //    "A" => this.Address.ToString(),
            //    "F" => HasFormula ? this.FormulaA1 : string.Empty,
            //    "NF" => Style.NumberFormat.Format,
            //    "FG" => Style.Font.FontColor.ToString(),
            //    "BG" => Style.Fill.BackgroundColor.ToString(),
            //    "V" => GetFormattedString(),
            //    _ => throw new FormatException($"Format {format} was not recognised."),
            //};
    /// <summary>
    /// Returns a string that represents the current state of the FsCell according to the format.
    /// </summary>
    // /// <param name="format">A: address, F: formula, NF: number format, BG: background color, FG: foreground color, V: formatted value</param>
    // /// <returns></returns>
    override self.ToString() = 
        let cellRef = CellReference.indexToColAdress (uint self.ColumnNumber)
        $"{cellRef}{self.RowNumber} : {self.Value} | {self.DataType}"

    /// <summary>
    /// Copies and replaces DataType and Value from a given FsCell into this one.
    /// </summary>
    member self.CopyFrom(otherCell : FsCell) = 
        self.DataType <- otherCell.DataType
        self.Value <- otherCell.Value

    /// <summary>
    /// Copies DataType and Value from this FsCell to a given one and replaces theirs.
    /// </summary>
    member self.CopyTo(target : FsCell) = 
        target.DataType <- self.DataType
        target.Value <- self.Value

    /// <summary>
    /// Copies and replaces DataType and Value from a source FsCell into a target FsCell. Returns the target cell.
    /// </summary>
    static member copyFromTo (sourceCell : FsCell) (targetCell : FsCell) =
        targetCell.DataType <- sourceCell.DataType
        targetCell.Value <- sourceCell.Value
        targetCell

    /// <summary>
    /// Creates a deep copy of this FsCell.
    /// </summary>
    member self.Copy() =
        let value = self.Value
        let dt = self.DataType
        let addr = self.Address.Copy()
        FsCell(value, dt, addr)

    /// <summary>
    /// Returns a deep copy of a given FsCell.
    /// </summary>
    static member copy (cell : FsCell) =
        cell.Copy()

    /// <summary>
    /// Gets the cell's value converted to the T type.
    /// <para>FsSpreadsheet will try to convert the current value to type 'T.</para>
    /// <para>An exception will be thrown if the current value cannot be converted to the T type.</para>
    /// </summary>
    /// <typeparam name="T">The return type.</typeparam>
    /// <exception cref="ArgumentException"></exception>
    member inline self.GetValueAs<'T>() =
        match (typeof<'T>) with
        | t when t = typeof<string>           -> self.Value |> box
        | t when t = typeof<bool>             -> bool.Parse (self.Value) |> box
        | t when t = typeof<float>            -> Double.Parse (self.Value) |> box
        | t when t = typeof<int>              -> Int32.Parse (self.Value) |> box
        
        | t when t = typeof<int16>            -> Int16.Parse (self.Value) |> box
        | t when t = typeof<int64>            -> Int64.Parse (self.Value) |> box
        
        | t when t = typeof<uint>             -> UInt32.Parse (self.Value) |> box
        | t when t = typeof<uint16>           -> UInt16.Parse (self.Value) |> box
        | t when t = typeof<uint64>           -> UInt64.Parse (self.Value) |> box

        | t when t = typeof<single>           -> Single.Parse (self.Value) |> box
        | t when t = typeof<decimal>          -> Decimal.Parse (self.Value) |> box
        | t when t = typeof<Guid>             -> Guid.Parse (self.Value) |> box
        | t when t = typeof<char>             -> Char.Parse (self.Value) |> box
        | t when t = typeof<DateTime>         -> DateTime.Parse (self.Value) |> box
        | t -> failwith $"FsCell with value {self.Value} cannot be parsed to {typeof<double>.Name}."
        
        :?> 'T

    /// <summary>
    /// Gets the FsCell's value converted to the 'T type.
    /// 
    /// FsSpreadsheet will try to convert the current value to type 'T. </summary>
    /// <typeparam name="T">The return type.</typeparam>
    /// <exception cref="System.ArgumentException">if the current value cannot be converted to the 'T type.</exception>
    static member inline getValueAs<'T>(cell : FsCell) =
        cell.GetValueAs<'T>()

    /// <summary>
    /// Sets the FsCell's value.
    /// <para>FsSpreadsheet will try to translate it to the corresponding type, if it cannot, the value will be left as a string.</para>
    /// </summary>
    /// <value>
    /// The object containing the value to set.
    /// </value>
    member self.SetValueAs<'T>(value) = 
        let t,v = DataType.InferCellValue value
        _dataType <- t
        _cellValue <- v

    /// <summary>
    /// Sets an FsCell's value.
    /// <para>FsSpreadsheet will try to translate it to the corresponding type, if it cannot, the value will be left as a string.</para>
    /// </summary>
    /// <value>
    /// The object containing the value to set.
    /// </value>
    static member setValueAs<'T> value (cell : FsCell)= 
        cell.SetValueAs<'T>(value)
        cell

    ///// <summary>
    ///// Sets the type of this FsCell's data.
    ///// <para>Changing the data type will cause FsSpreadsheet to convert the current value to the new DataType.</para>
    ///// <para>An exception will be thrown if the current value cannot be converted to the new DataType.</para>
    ///// </summary>
    ///// <param name="dataType">Type of the data.</param>
    ///// <returns></returns>
    //member self.SetDataType(dataType) = 
    //    self.DataType <- dataType

    ///// <summary>
    ///// Sets the type of an FsCell's data.
    ///// <para>Changing the data type will cause FsSpreadsheet to convert the current value to the new DataType.</para>
    ///// <para>An exception will be thrown if the current value cannot be converted to the new DataType.</para>
    ///// </summary>
    ///// <param name="dataType">Type of the data.</param>
    ///// <returns></returns>
    //static member setDataType dataType (cell : FsCell) =
    //    cell.DataType <- dataType
    //    cell


//################################################################################
//################################################################################
// Not implemented yet
//################################################################################

    ///// Gets or sets the FsCell's associated FsWorksheet.
    //member self.Worksheet = raise (System.NotImplementedException())


    //member internal self.SharedStringId = raise (System.NotImplementedException())

    //member self.Active = raise (System.NotImplementedException())
    


    ///// <summary>
    ///// Calculated value of cell formula. Is used for decreasing number of computations perfromed.
    ///// May hold invalid value when <see cref="NeedsRecalculation"/> flag is True.
    ///// </summary>
    //member self.CachedValue = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Returns the current region. The current region is a range bounded by any combination of blank rows and blank columns
    ///// </summary>
    ///// <value>
    ///// The current region.
    ///// </value>
    //member self.CurrentRegion = raise (System.NotImplementedException())
    

    
    ///// <summary>
    ///// Gets or sets the cell's formula with A1 references.
    ///// </summary>
    ///// <value>The formula with A1 references.</value>
    //member self.FormulaA1 = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Gets or sets the cell's formula with R1C1 references.
    ///// </summary>
    ///// <value>The formula with R1C1 references.</value>
    //member self.FormulaR1C1 = raise (System.NotImplementedException())
    
    //member self.FormulaReference = raise (System.NotImplementedException())
    
    //member self.HasArrayFormula = raise (System.NotImplementedException())
    
    //member self.HasComment = raise (System.NotImplementedException())
    
    //member self.HasDataValidation = raise (System.NotImplementedException())
    
    //member self.HasFormula = raise (System.NotImplementedException())
    
    //member self.HasHyperlink = raise (System.NotImplementedException())
    
    //member self.HasRichText = raise (System.NotImplementedException())
    
    //member self.HasSparkline = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Flag indicating that previously calculated FsCell value may be not valid anymore and has to be re-evaluated.
    ///// </summary>
    //member self.NeedsRecalculation = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Gets or sets a value indicating whether this cell's text should be shared or not.
    ///// </summary>
    ///// <value>
    /////   If false the cell's text will not be shared and stored as an inline value.
    ///// </value>
    //member self.ShareString = raise (System.NotImplementedException())
    
    //member self.Sparkline = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Gets or sets the FsCell's style.
    ///// </summary>
    //member self.Style = raise (System.NotImplementedException())
    



    //// ------------------
    //// NON-STATIC METHODS
    //// ------------------
    
    //member self.AddConditionalFormat()  = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates a named range out of this cell.
    ///// <para>If the named range exists, it will add this range to that named range.</para>
    ///// <para>The default scope for the named range is Workbook.</para>
    ///// </summary>
    ///// <param name="rangeName">Name of the range.</param>
    //member self.AddToNamed(rangeName)  = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates a named range out of this cell.
    ///// <para>If the named range exists, it will add this range to that named range.</para>
    ///// <param name="rangeName">Name of the range.</param>
    ///// <param name="scope">The scope for the named range.</param>
    ///// </summary>
    //member self.AddToNamed(rangeName, scope) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates a named range out of this cell.
    ///// <para>If the named range exists, it will add this range to that named range.</para>
    ///// <param name="rangeName">Name of the range.</param>
    ///// <param name="scope">The scope for the named range.</param>
    ///// <param name="comment">The comments for the named range.</param>
    ///// </summary>
    //member self.AddToNamed(rangeName, scope, comment) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Returns this cell as an IXLRange.
    ///// </summary>
    //member self.AsRange()  = raise (System.NotImplementedException())
    
    //member self.CellAbove() = raise (System.NotImplementedException())
    
    //member self.CellAbove(step) = raise (System.NotImplementedException())
    
    //member self.CellBelow() = raise (System.NotImplementedException())
    
    //member self.CellBelow(step) = raise (System.NotImplementedException())
    
    //member self.CellLeft() = raise (System.NotImplementedException())
    
    //member self.CellLeft(step) = raise (System.NotImplementedException())
    
    //member self.CellRight() = raise (System.NotImplementedException())
    
    //member self.CellRight(step) = raise (System.NotImplementedException())
    
    //// see https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Excel/Cells/XLCell.cs#L860
    ///// <summary>
    ///// Clears the contents of this FsCell.
    ///// </summary>
    ///// <param name="clearOptions">Specify what you want to clear.</param>
    //member self.Clear(clearOptions(* = XLClearOptions.All*)) = raise (System.NotImplementedException())
    
    ////member self.CopyFrom(member self.otherCell);
    

    ///// <summary>
    ///// Creates a new comment for the cell, replacing the existing one.
    ///// </summary>
    //member self.CreateComment() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates a new data validation rule for the cell, replacing the existing one.
    ///// </summary>
    //member self.CreateDataValidation() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates a new hyperlink replacing the existing one.
    ///// </summary>
    //member self.CreateHyperlink() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Replaces a value of the cell with a newly created rich text object.
    ///// </summary>
    //member self.CreateRichText() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Deletes the current cell and shifts the surrounding cells according to the shiftDeleteCells parameter.
    ///// </summary>
    ///// <param name="shiftDeleteCells">How to shift the surrounding cells.</param>
    //member self.Delete(shiftDeleteCells) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Gets the cell's value converted to Boolean.
    ///// <para>ClosedXML will try to covert the current value to Boolean.</para>
    ///// <para>An exception will be thrown if the current value cannot be converted to Boolean.</para>
    ///// </summary>
    //member self.GetBoolean() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Returns the comment for the cell or create a new instance if there is no comment on the cell.
    ///// </summary>
    //member self.GetComment() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Returns a data validation rule assigned to the cell, if any, or creates a new instance of data validation rule if no rule exists.
    ///// </summary>
    //member self.GetDataValidation() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Gets the cell's value converted to DateTime.
    ///// <para>ClosedXML will try to covert the current value to DateTime.</para>
    ///// <para>An exception will be thrown if the current value cannot be converted to DateTime.</para>
    ///// </summary>
    //member self.GetDateTime() = raise (System.NotImplementedException())
    
   
    ///// <summary>
    ///// Gets the cell's value formatted depending on the cell's data type and style.
    ///// </summary>
    //member self.GetFormattedString() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Returns a hyperlink for the cell, if any, or creates a new instance is there is no hyperlink.
    ///// </summary>
    //member self.GetHyperlink() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Returns the value of the cell if it formatted as a rich text.
    ///// </summary>
    //member self.GetRichText() = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Gets the FsCell's value converted to a String.
    ///// </summary>
    //member self.GetString() = value
    
    ///// <summary>
    ///// Gets the cell's value converted to TimeSpan.
    ///// <para>ClosedXML will try to covert the current value to TimeSpan.</para>
    ///// <para>An exception will be thrown if the current value cannot be converted to TimeSpan.</para>
    ///// </summary>
    //member self.GetTimeSpan() = raise (System.NotImplementedException())
    


    //member self.InsertCellsAbove(numberOfRows) = raise (System.NotImplementedException())
    
    //member self.InsertCellsAfter(numberOfColumns) = raise (System.NotImplementedException())
    
    //member self.InsertCellsBefore(numberOfColumns) = raise (System.NotImplementedException())
    
    //member self.InsertCellsBelow(numberOfRows)  = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the IEnumerable data elements and returns the range it occupies.
    ///// </summary>
    ///// <param name="data">The IEnumerable data.</param>
    //member self.InsertData(data)  = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the IEnumerable data elements and returns the range it occupies.
    ///// </summary>
    ///// <param name="data">The IEnumerable data.</param>
    ///// <param name="transpose">if set to <c>true</c> the data will be transposed before inserting.</param>
    ///// <returns></returns>
    //member self.InsertData(data, transpose) = raise (System.NotImplementedException())
    
    /////// <summary>
    /////// Inserts the data of a data table.
    /////// </summary>
    /////// <param name="dataTable">The data table.</param>
    /////// <returns>The range occupied by the inserted data</returns>
    ////member self.InsertData(dataTable) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the IEnumerable data elements as a table and returns it.
    ///// <para>The new table will receive a generic name: Table#</para>
    ///// </summary>
    ///// <param name="data">The table data.</param>
    //member self.InsertTable<'T>(data) = raise (System.NotImplementedException())
    
    /////// <summary>
    /////// Inserts the IEnumerable data elements as a table and returns it.
    /////// <para>The new table will receive a generic name: Table#</para>
    /////// </summary>
    /////// <param name="data">The table data.</param>
    /////// <param name="createTable">
    /////// if set to <c>true</c> it will create an Excel table.
    /////// <para>if set to <c>false</c> the table will be created in memory.</para>
    /////// </param>
    ////member self.InsertTable<'T>(data, createTable)  = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates an Excel table from the given IEnumerable data elements.
    ///// </summary>
    ///// <param name="data">The table data.</param>
    ///// <param name="tableName">Name of the table.</param>
    //member self.InsertTable<'T>(data, tableName) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the IEnumerable data elements as a table and returns it.
    ///// </summary>
    ///// <param name="data">The table data.</param>
    ///// <param name="tableName">Name of the table.</param>
    ///// <param name="createTable">
    ///// if set to <c>true</c> it will create an Excel table.
    ///// <para>if set to <c>false</c> the table will be created in memory.</para>
    ///// </param>
    //member self.InsertTable<'T>(data, tableName, createTable) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the DataTable data elements as a table and returns it.
    ///// <para>The new table will receive a generic name: Table#</para>
    ///// </summary>
    ///// <param name="data">The table data.</param>
    //member self.InsertTable(data) = raise (System.NotImplementedException())
    
    /////// <summary>
    /////// Inserts the DataTable data elements as a table and returns it.
    /////// <para>The new table will receive a generic name: Table#</para>
    /////// </summary>
    /////// <param name="data">The table data.</param>
    /////// <param name="createTable">
    /////// if set to <c>true</c> it will create an Excel table.
    /////// <para>if set to <c>false</c> the table will be created in memory.</para>
    /////// </param>
    ////member self.InsertTable(data, createTable) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Creates an Excel table from the given DataTable data elements.
    ///// </summary>
    ///// <param name="data">The table data.</param>
    ///// <param name="tableName">Name of the table.</param>
    //member self.InsertTable(data, tableName)  = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Inserts the DataTable data elements as a table and returns it.
    ///// </summary>
    ///// <param name="data">The table data.</param>
    ///// <param name="tableName">Name of the table.</param>
    ///// <param name="createTable">
    ///// if set to <c>true</c> it will create an Excel table.
    ///// <para>if set to <c>false</c> the table will be created in memory.</para>
    ///// </param>
    //member self.InsertTable(data, tableName, createTable) = raise (System.NotImplementedException())
    
    ///// <summary>
    ///// Invalidate <see cref="CachedValue"/> so the formula will be re-evaluated next time <see cref="Value"/> is accessed.
    ///// If cell does not contain formula nothing happens.
    ///// </summary>
    //member self.InvalidateFormula() = raise (System.NotImplementedException())
    
    //member self.IsEmpty() = raise (System.NotImplementedException())
    
    //[<System.Obsolete("Use the overload with XLCellsUsedOptions")>]
    //member self.IsEmpty(includeFormats) = raise (System.NotImplementedException())
    
    ////member self.IsEmpty(options) = raise (System.NotImplementedException())
    
    //member self.IsMerged() = raise (System.NotImplementedException())
    
    //member self.MergedRange() = raise (System.NotImplementedException())
    
    //member self.Select() = raise (System.NotImplementedException())
    
    //member self.SetActive(value(* = true*)) = raise (System.NotImplementedException())
    

    
    //[<System.Obsolete("Use GetDataValidation to access the existing rule, or CreateDataValidation() to create a new one.")>]
    //member self.SetDataValidation() = raise (System.NotImplementedException())
    
    //member self.SetFormulaA1(formula) = raise (System.NotImplementedException())
    
    //member self.SetFormulaR1C1(formula) = raise (System.NotImplementedException())
    
    //member self.SetHyperlink(hyperlink) = raise (System.NotImplementedException())
    

    //member self.TableCellType() = raise (System.NotImplementedException())
    

    
    //// same problem like with .GetValue<'T>
    //member self.TryGetValue<'T>(value) = raise (System.NotImplementedException())