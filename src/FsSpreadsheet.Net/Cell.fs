﻿namespace FsSpreadsheet.Net

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet

open FsSpreadsheet

/// <summary>
/// Functions for creating and manipulating Cells.
/// </summary>
module Cell =


    /// <summary>
    /// Functions for manipulating CellValues.
    /// </summary>
    module CellValue = 

        /// <summary>
        /// Creates an empty CellValue.
        /// </summary>
        let empty() = CellValue()

        /// <summary>
        /// Create a new CellValue containing the given string.
        /// </summary>
        let create (value : string) = CellValue(value)

        /// <summary>
        /// Returns the value stored inside the CellValue.
        /// </summary>
        let getValue (cellValue : CellValue) = cellValue.Text

        /// <summary>
        /// Sets the value inside the CellValue.
        /// </summary>
        let setValue (value : string) (cellValue : CellValue) =  cellValue.Text <- value

    /// <summary>
    /// Takes a CellValue and returns the appropriate DataType.
    /// </summary>
    /// <remarks>DataType is the FsSpreadsheet representation of the CellValue enum in OpenXml.</remarks>
    let cellValuesToDataType (cellValue : CellValues) =
        match cellValue with
        | CellValues.SharedString
        | CellValues.InlineString
        | CellValues.String     -> String
        | CellValues.Boolean    -> Boolean
        | CellValues.Number     -> Number
        | CellValues.Date       -> Date
        | CellValues.Error      -> Empty
        | _                     -> failwith $"cellValue {cellValue.ToString()} can not be transferred to DataType."

    /// <summary>
    /// Creates an empty Cell.
    /// </summary>
    let empty () = Cell()

    /// <summary>
    /// Returns the proper CellValues case for the given value.
    /// </summary>
    let inferCellValue (value : 'T) = 
        let value = box value
        match value with
        | :? char as c -> CellValues.String,c.ToString()
        | :? string as s -> CellValues.String,s.ToString()
        | :? bool as c -> CellValues.Boolean,c.ToString()
        | :? byte as i -> CellValues.Number,i.ToString()
        | :? sbyte as i -> CellValues.Number,i.ToString()
        | :? int as i -> CellValues.Number,i.ToString()
        | :? int16 as i -> CellValues.Number,i.ToString()
        | :? int64 as i -> CellValues.Number,i.ToString()
        | :? uint as i -> CellValues.Number,i.ToString()
        | :? uint16 as i -> CellValues.Number,i.ToString()
        | :? uint64 as i -> CellValues.Number,i.ToString()
        | :? single as i -> CellValues.Number,i.ToString()
        | :? float as i -> CellValues.Number,i.ToString()
        | :? decimal as i -> CellValues.Number,i.ToString()
        | :? System.DateTime as d -> CellValues.Date,d.Date.ToString()
        | _ ->  CellValues.String,value.ToString()

    /// <summary>
    /// Creates a Cell from a CellValues type case, a "A1" style reference, and a CellValue containing the value string.
    /// </summary>
    let create (dataType : CellValues option) (reference : string) (value : CellValue) = 
        match dataType with
        | Some dataType -> Cell(CellReference = StringValue.FromString reference, DataType = EnumValue(dataType), CellValue = value)
        | None -> Cell(CellReference = StringValue.FromString reference, CellValue = value)

    /// <summary>
    /// Creates a Cell from a CellValues type case, a "A1" style reference, and a CellValue containing the value string.
    /// </summary>
    let createWithFormat doc (dataType : CellValues option) (reference : string) (cellFormat : CellFormat) (value : CellValue) = 
        let styleSheet = Stylesheet.getOrInit doc
        let i = Stylesheet.CellFormat.appendOrGetIndex cellFormat styleSheet
        match dataType with
        | Some dataType -> Cell(StyleIndex = UInt32Value(uint32 i),CellReference = StringValue.FromString reference, DataType = EnumValue(dataType), CellValue = value)
        | None -> Cell(StyleIndex = UInt32Value(uint32 i),CellReference = StringValue.FromString reference, CellValue = value)

    /// <summary>
    /// Sets the preserve attribute of a Cell.
    /// </summary>
    let setSpacePreserveAttribute (c : Cell) =
        c.SetAttribute(OpenXmlAttribute("xml:space","","preserve"))
        c

    /// <summary>
    /// Create a cell using a shared string table, also returns the updated shared string table.
    /// </summary>
    let fromValue (sharedStringTable : SharedStringTable Option) columnIndex rowIndex (value : 'T) = 
        let value = box value
        match value with
        | :? string as s when sharedStringTable.IsSome-> 
            let sharedStringTable = sharedStringTable.Value
            let reference = CellReference.ofIndices columnIndex (rowIndex)
            match SharedStringTable.tryGetIndexByString s sharedStringTable with
            | Some i -> 
                i
                |> string
                |> CellValue.create
                |> create (Some CellValues.SharedString) reference
            | None ->
                let updatedSharedStringTable = 
                    sharedStringTable
                    |> SharedStringTable.SharedStringItem.add (SharedStringTable.SharedStringItem.create s) 

                updatedSharedStringTable
                |> SharedStringTable.count
                |> string
                |> CellValue.create
                |> create (Some CellValues.SharedString) reference 
            |> fun c -> 
                if s.EndsWith " " then
                    setSpacePreserveAttribute c
                else c

        | _  -> 
            let valType,value = inferCellValue value
            let reference = CellReference.ofIndices columnIndex (rowIndex)
            create (Some valType) reference (CellValue.create value)
            |> fun c ->
                if value.EndsWith " " then
                    setSpacePreserveAttribute c
                else c

    let getCellContent (doc : Packaging.SpreadsheetDocument) (value : string) (dataType : DataType) = 
        let sharedStringTable = SharedStringTable.tryGet doc
        match dataType with
        | DataType.String when sharedStringTable.IsSome-> 
            let sharedStringTable = sharedStringTable.Value
            match SharedStringTable.tryGetIndexByString value sharedStringTable with
            | Some i -> 
                i
                |> string
            | None ->
                let updatedSharedStringTable = 
                    sharedStringTable
                    |> SharedStringTable.SharedStringItem.add (SharedStringTable.SharedStringItem.create value) 

                updatedSharedStringTable
                |> SharedStringTable.count
                |> string
            |> fun v -> {|DataType = Some CellValues.SharedString; Value = v; Format = None|}
        | DataType.String ->    
            {|DataType = Some CellValues.String; Value = value; Format = None|}
        | DataType.Boolean ->
            {|DataType = Some CellValues.Boolean; Value = System.Boolean.Parse value |> FsCellAux.boolConverter; Format = None|}
        | DataType.Number ->
            {|DataType = Some CellValues.Number; Value = value; Format = None|}
        | DataType.Date ->      
            //let cellFormat = CellFormat(NumberFormatId = UInt32Value 19u, ApplyNumberFormat = BooleanValue true)
            let value = System.DateTime.Parse(value).ToOADate() |> string
            let cellFormat = 
                if value.Contains(".") then
                    Stylesheet.CellFormat.getDefaultDateTime()
                else 
                    Stylesheet.CellFormat.getDefaultDate()
            {|DataType = None; Value = value; Format = Some cellFormat|}
        | DataType.Empty ->     {|DataType = None; Value = value; Format = None|}

    /// <summary>
    /// Create a cell using a shared string table, also returns the updated shared string table.
    /// </summary>
    let fromValueWithDataType (doc : Packaging.SpreadsheetDocument) columnIndex rowIndex (value : string) (dataType : DataType) = 
        let reference = CellReference.ofIndices columnIndex (rowIndex)
        let cellContent = getCellContent doc value dataType
        if cellContent.Format.IsSome then
            createWithFormat doc cellContent.DataType reference cellContent.Format.Value (CellValue.create cellContent.Value)
        else
            create cellContent.DataType reference (CellValue.create cellContent.Value)
        |> fun c -> 
                if value.EndsWith " " then
                    setSpacePreserveAttribute c
                else c
    /// <summary>
    /// Gets "A1"-style Cell reference.
    /// </summary>
    let getReference (cell : Cell) = cell.CellReference.Value

    /// <summary>
    /// Sets "A1"-style Cell reference.
    /// </summary>
    let setReference (reference) (cell : Cell) = 
        cell.CellReference <- StringValue.FromString reference
        cell

    /// <summary>
    /// Gets Some type if existent. Else returns None.
    /// </summary>
    let tryGetType (cell : Cell) = 
        if cell.DataType <> null then
            Some cell.DataType.Value
        else
            None
    
    /// <summary>
    /// Gets a Cell type.
    /// </summary>
    let getType (cell : Cell) = cell.DataType.Value

    /// <summary>
    /// Sets a Cell type.
    /// </summary>
    let setType (dataType : CellValues) (cell : Cell) = 
        cell.DataType <- EnumValue(dataType)
        cell

    /// <summary>
    /// Gets Some CellValue if cellValue is existent. Else returns None.
    /// </summary>
    let tryGetCellValue (cell : Cell) = 
        if cell.CellValue <> null then
            Some cell.CellValue
        else
            None

    /// <summary>
    /// Gets the CellValue.
    /// </summary>
    let getCellValue (cell : Cell) = cell.CellValue
    
    let tryGetInnerText (cell : Cell) = 
        if cell.HasChildren then
            Some cell.InnerText
        else 
            None

    /// <summary>
    /// Maps a Cell to the value string using a shared string table.
    /// </summary>
    let tryGetValue (sharedStringTable:SST Option) (cell:Cell) =
        match cell |> tryGetType with
        | Some (CellValues.SharedString) when sharedStringTable.IsSome->
            let sharedStringTable = sharedStringTable.Value
            cell
            |> tryGetCellValue
            |> Option.map (
                CellValue.getValue 
                >> int
                >> fun i -> sharedStringTable.[i]          
            )
        | Some CellValues.InlineString ->
            cell
            |> tryGetInnerText
        | _ ->
            cell
            |> tryGetCellValue
            |> Option. map CellValue.getValue   

    /// <summary>
    /// Maps a Cell to the value string using a sharedStringTable.
    /// </summary>
    let getValue (sharedStringTable : SST Option) (cell : Cell) =
        match cell |> tryGetType with
        | Some (CellValues.SharedString) when sharedStringTable.IsSome ->
            let sharedStringTable = sharedStringTable.Value
            let sharedStringTableIndex = 
                cell
                |> getCellValue
                |> CellValue.getValue
                |> int

            sharedStringTable.[sharedStringTableIndex]
        | _ ->
            cell
            |> getCellValue
            |> CellValue.getValue   

    /// <summary>
    /// Sets a CellValue.
    /// </summary>
    let setValue (value : CellValue) (cell : Cell) = 
        cell.CellValue <- value
        cell

    /// <summary>
    /// Includes a value from the sharedStringTable in Cell.CellValue.Text.
    /// </summary>
    let includeSharedStringValue (sharedStringTable:SST) (cell:Cell) =
        if not (isNull cell.DataType) then  
            match cell |> tryGetType with
            | Some (CellValues.SharedString) ->
                let index = int cell.InnerText
                match sharedStringTable |> Array.tryItem index with 
                | Some value -> 
                    cell.DataType <- EnumValue(CellValues.String)
                    cell.CellValue.Text <- value
                | None ->
                    cell.CellValue.Text <- cell.InnerText
                cell  

            | _ -> cell
        else
            try cell.CellValue.Text <- cell.InnerText
            with _ -> cell.CellValue <- CellValue.empty()
            cell
