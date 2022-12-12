namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet

open FsSpreadsheet

/// Functions for creating and manipulating Cells.
module Cell =


    /// Functions for manipulating CellValues.
    module CellValue = 

        /// Creates an empty CellValue.
        let empty() = CellValue()

        /// Create a new cellValue containing the given string.
        let create (value : string) = CellValue(value)

        /// Returns the value stored inside the CellValue.
        let getValue (cellValue : CellValue) = cellValue.Text

        /// Sets the value inside the CellValue.
        let setValue (value : string) (cellValue : CellValue) =  cellValue.Text <- value

    let cellValuesFromDataType (dataType : DataType) =
        match dataType with
        | String    -> CellValues.String
        | Boolean   -> CellValues.Boolean
        | Number    -> CellValues.Number
        | Date      -> CellValues.Date
        | Empty     -> CellValues.Error

    /// Creates an empty cell.
    let empty () = Cell()

    /// Returns the proper CellValues case for the given value.
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

    /// Creates a cell from a CellValues type case, a "A1" style reference, and a CellValue containing the value string.
    let create (dataType : CellValues) (reference : string) (value : CellValue) = 
        Cell(CellReference = StringValue.FromString reference, DataType = EnumValue(dataType), CellValue = value)
        
    let setSpacePreserveAttribute (c : Cell) =
        c.SetAttribute(OpenXmlAttribute("xml:space","","preserve"))
        c

    /// Create a cell using a shared string table, also returns the updated shared string table.
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
                |> create CellValues.SharedString reference
            | None ->
                let updatedSharedStringTable = 
                    sharedStringTable
                    |> SharedStringTable.SharedStringItem.add (SharedStringTable.SharedStringItem.create s) 

                updatedSharedStringTable
                |> SharedStringTable.count
                |> string
                |> CellValue.create
                |> create CellValues.SharedString reference 
            |> fun c -> 
                if s.EndsWith " " then
                    setSpacePreserveAttribute c
                else c

        | _  -> 
            let valType,value = inferCellValue value
            let reference = CellReference.ofIndices columnIndex (rowIndex)
            create valType reference (CellValue.create value)
            |> fun c ->
                if value.EndsWith " " then
                    setSpacePreserveAttribute c
                else c

    /// Create a cell using a shared string table, also returns the updated shared string table.
    let fromValueWithDataType (sharedStringTable : SharedStringTable Option) columnIndex rowIndex (value : string) (dataType : DataType) = 
        match dataType with
        | DataType.String when sharedStringTable.IsSome-> 
            let sharedStringTable = sharedStringTable.Value
            let reference = CellReference.ofIndices columnIndex (rowIndex)
            match SharedStringTable.tryGetIndexByString value sharedStringTable with
            | Some i -> 
                i
                |> string
                |> CellValue.create
                |> create CellValues.SharedString reference
            | None ->
                let updatedSharedStringTable = 
                    sharedStringTable
                    |> SharedStringTable.SharedStringItem.add (SharedStringTable.SharedStringItem.create value) 

                updatedSharedStringTable
                |> SharedStringTable.count
                |> string
                |> CellValue.create
                |> create CellValues.SharedString reference 
            |> fun c -> 
                if value.EndsWith " " then
                    setSpacePreserveAttribute c
                else c

        | _  -> 
           let valType = cellValuesFromDataType dataType
           let reference = CellReference.ofIndices columnIndex (rowIndex)
           create valType reference (CellValue.create value)
           |> fun c ->
                if value.EndsWith " " then
                    setSpacePreserveAttribute c
                else c

    /// Gets "A1"-style cell reference.
    let getReference (cell : Cell) = cell.CellReference.Value

    /// Sets "A1"-style cell reference.
    let setReference (reference) (cell : Cell) = 
        cell.CellReference <- StringValue.FromString reference
        cell

    /// Gets Some type if existent. Else returns None.
    let tryGetType (cell : Cell) = 
        if cell.DataType <> null then
            Some cell.DataType.Value
        else
            None
    
    /// Gets a Cell type.
    let getType (cell : Cell) = cell.DataType.Value

    /// Sets a Cell type.
    let setType (dataType : CellValues) (cell : Cell) = 
        cell.DataType <- EnumValue(dataType)
        cell

    /// Gets Some CellValue if cellValue is existent. Else returns None.
    let tryGetCellValue (cell : Cell) = 
        if cell.CellValue <> null then
            Some cell.CellValue
        else
            None

    /// Gets the CellValue.
    let getCellValue (cell : Cell) = cell.CellValue
    
    /// Maps a cell to the value string using a shared string table.
    let tryGetValue (sharedStringTable:SharedStringTable Option) (cell:Cell) =
        match cell |> tryGetType with
        | Some (CellValues.SharedString) when sharedStringTable.IsSome->
            let sharedStringTable = sharedStringTable.Value
            cell
            |> tryGetCellValue
            |> Option.map (
                CellValue.getValue 
                >> int
                >> fun i -> SharedStringTable.getText i sharedStringTable
                >> SharedStringTable.SharedStringItem.getText                   
            )
    
        | _ ->
            cell
            |> tryGetCellValue
            |> Option. map CellValue.getValue   



    /// Maps a Cell to the value string using a sharedStringTable.
    let getValue (sharedStringTable : SharedStringTable Option) (cell : Cell) =
        match cell |> tryGetType with
        | Some (CellValues.SharedString) when sharedStringTable.IsSome->
            let sharedStringTable = sharedStringTable.Value

            let sharedStringTableIndex = 
                cell
                |> getCellValue
                |> CellValue.getValue
                |> int

            sharedStringTable
            |> SharedStringTable.getText sharedStringTableIndex
            |> SharedStringTable.SharedStringItem.getText
        | _ ->
            cell
            |> getCellValue
            |> CellValue.getValue   

    /// Sets a CellValue.
    let setValue (value : CellValue) (cell : Cell) = 
        cell.CellValue <- value
        cell

    /// Includes a value from the sharedStringTable in Cell.CellValue.Text.
    let includeSharedStringValue (sharedStringTable:SharedStringTable) (cell:Cell) =
        if not (isNull cell.DataType) then  
            match cell |> tryGetType with
            | Some (CellValues.SharedString) ->
                let index = int cell.InnerText
                match sharedStringTable |> Seq.tryItem index with 
                | Some value -> 
                    cell.DataType <- EnumValue(CellValues.String)
                    cell.CellValue.Text <- value.InnerText
                | None ->
                    cell.CellValue.Text <- cell.InnerText
                cell  

            | _ -> cell
        else        
            cell.CellValue.Text <- cell.InnerText
            cell
