namespace FsSpreadsheet

open System

open Fable.Core

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

[<AttachMembers>]
type FsCell (value : IConvertible, ?dataType : DataType, ?address : FsAddress) =
    
    // TODO: Maybe save as IConvertible
    let mutable _cellValue = string value
    let mutable _dataType = dataType |> Option.defaultValue DataType.String
    let mutable _comment  = ""
    let mutable _hyperlink = ""
    let mutable _richText = ""
    let mutable _formulaA1 = ""
    let mutable _formulaR1C1 = ""

    let mutable _rowIndex : int = address |> Option.map (fun a -> a.RowNumber) |> Option.defaultValue 0
    let mutable _columnIndex : int = address |> Option.map (fun a -> a.ColumnNumber) |> Option.defaultValue 0


    /// Creates an empty FsCell, set at row 0, column 0 (1-based).
    static member empty () = FsCell ("", DataType.Empty, FsAddress(0,0))

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
    /// Gets the value as string
    /// </summary>
    member self.ValueAsString() =
        self.Value

    /// <summary>
    /// Gets the value as string
    /// </summary>
    static member getValueAsString (cell : FsCell) =
        cell.ValueAsString()

    /// <summary>
    /// Gets the value as bool
    /// </summary>
    member self.ValueAsBool() =
        bool.Parse (self.Value)

    /// <summary>
    /// Gets the value as bool
    /// </summary>
    static member getValueAsBool (cell : FsCell) =
        cell.ValueAsBool()

    /// <summary>
    /// Gets the value as float
    /// </summary>
    member self.ValueAsFloat() =
        Double.Parse (self.Value)

    /// <summary>
    /// Gets the value as float
    /// </summary>
    static member getValueAsFloat (cell : FsCell) =
        cell.ValueAsFloat()

    /// <summary>
    /// Gets the value as int
    /// </summary>
    member self.ValueAsInt() =
        Int32.Parse (self.Value)

    /// <summary>
    /// Gets the value as int
    /// </summary>
    static member getValueAsInt (cell : FsCell) =
        cell.ValueAsInt()

    /// <summary>
    /// Gets the value as uint
    /// </summary>
    member self.ValueAsUInt() =
        UInt32.Parse (self.Value)

    /// <summary>
    /// Gets the value as uint
    /// </summary>
    static member getValueAsUInt (cell : FsCell) =
        cell.ValueAsUInt()

    /// <summary>
    /// Gets the value as long
    /// </summary>
    member self.ValueAsLong() =
        Int64.Parse (self.Value)

    /// <summary>
    /// Gets the value as long
    /// </summary>
    static member getValueAsLong (cell : FsCell) =
        cell.ValueAsLong()

    /// <summary>
    /// Gets the value as ulong
    /// </summary>
    member self.ValueAsULong() =
        UInt64.Parse (self.Value)

    /// <summary>
    /// Gets the value as ulong
    /// </summary>
    static member getValueAsULong (cell : FsCell) =
        cell.ValueAsULong()

    /// <summary>
    /// Gets the value as double
    /// </summary>
    member self.ValueAsDouble() =
        Double.Parse (self.Value)

    /// <summary>
    /// Gets the value as double
    /// </summary>
    static member getValueAsDouble (cell : FsCell) =
        cell.ValueAsDouble()

    /// <summary>
    /// Gets the value as decimal
    /// </summary>
    member self.ValueAsDecimal() =
        Decimal.Parse (self.Value)

    /// <summary>
    /// Gets the value as decimal
    /// </summary>
    static member getValueAsDecimal (cell : FsCell) =
        cell.ValueAsDecimal()

    /// <summary>
    /// Gets the value as DateTime
    /// </summary>
    member self.ValueAsDateTime() =
        DateTime.Parse (self.Value)

    /// <summary>
    /// Gets the value as DateTime
    /// </summary>
    static member getValueAsDateTime (cell : FsCell) =
        cell.ValueAsDateTime()

    /// <summary>
    /// Gets the value as Guid
    /// </summary>
    member self.ValueAsGuid() =
        Guid.Parse (self.Value)

    /// <summary>
    /// Gets the value as Guid
    /// </summary>
    static member getValueAsGuid (cell : FsCell) =
        cell.ValueAsGuid()

    /// <summary>
    /// Gets the value as char
    /// </summary>
    member self.ValueAsChar() =
        Char.Parse (self.Value)

    /// <summary>
    /// Gets the value as char
    /// </summary>
    static member getValueAsChar (cell : FsCell) =
        cell.ValueAsChar()

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
