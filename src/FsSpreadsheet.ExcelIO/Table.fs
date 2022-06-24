namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml

open FsSpreadsheet

/// Functions for working with tables. 
///
/// The table object itself just stores the name, the headers and the area in which the table lies.
/// The values are stored in the sheetData object associated with the same worksheet as the table.
///
/// Therefore, in order to work with tables, one should retrieve both the sheetData and the table. The value retrieval functions ask for both.
module Table =
    
    //  Helper functions for working with "A1:A1"-style table areas.
    /// The areas marks the area in which the table lies. 
    module Area =

        /// Given A1-based top left start and bottom right end indices, returns a "A1:A1"-style area-
        let ofBoundaries fromCellReference toCellReference = 
            sprintf "%s:%s" fromCellReference toCellReference
            |> StringValue.FromString

        /// Given a "A1:A1"-style area, returns A1-based cell start and end cellReferences.
        let toBoundaries (area : StringValue) = 
            area.Value.Split ':'
            |> fun a -> a.[0], a.[1]

        /// Gets the right boundary of the area.
        let rightBoundary (area : StringValue) = 
            toBoundaries area
            |> snd
            |> CellReference.toIndices
            |> fst

        /// Gets the left boundary of the area.
        let leftBoundary (area : StringValue) = 
            toBoundaries area
            |> fst
            |> CellReference.toIndices
            |> fst

        /// Gets the Upper boundary of the area.
        let upperBoundary (area : StringValue) = 
            toBoundaries area
            |> fst
            |> CellReference.toIndices
            |> snd

        /// Gets the lower boundary of the area.
        let lowerBoundary (area : StringValue) = 
            toBoundaries area
            |> snd
            |> CellReference.toIndices
            |> snd

        /// Moves both start and end of the area by the given amount (positive amount moves area to right and vice versa).
        let moveHorizontal amount (area : StringValue) =
            area
            |> toBoundaries
            |> fun (f,t) -> CellReference.moveHorizontal amount f, CellReference.moveHorizontal amount t
            ||> ofBoundaries

        /// Moves both start and end of the area by the given amount (positive amount moves area to right and vice versa).
        let moveVertical amount (area : StringValue) =
            area
            |> toBoundaries
            |> fun (f,t) -> CellReference.moveHorizontal amount f, CellReference.moveHorizontal amount t
            ||> ofBoundaries

        /// Extends the right boundary of the area by the given amount (positive amount increases area to right and vice versa).
        let extendRight amount (area : StringValue) =
            area
            |> toBoundaries
            |> fun (f,t) -> f, CellReference.moveHorizontal amount t
            ||> ofBoundaries

        /// Extends the left boundary of the area by the given amount (positive amount decreases the area to left and vice versa).
        let extendLeft amount (area : StringValue) =
            area
            |> toBoundaries
            |> fun (f,t) -> CellReference.moveHorizontal amount f, t
            ||> ofBoundaries

        /// Returns true if the column index of the reference exceeds the right boundary of the area.
        let referenceExceedsAreaRight reference area = 
            (reference |> CellReference.toIndices |> fst) 
                > (area |> rightBoundary)
        
        /// Returns true if the column index of the reference exceeds the left boundary of the area.
        let referenceExceedsAreaLeft reference area = 
            (reference |> CellReference.toIndices |> fst) 
                < (area |> leftBoundary)  
     
        /// Returns true if the column index of the reference exceeds the upper boundary of the area.
        let referenceExceedsAreaAbove reference area = 
            (reference |> CellReference.toIndices |> snd) 
                > (area |> upperBoundary)
        
        /// Returns true if the column index of the reference exceeds the lower boundary of the area.
        let referenceExceedsAreaBelow reference area = 
            (reference |> CellReference.toIndices |> snd) 
                < (area |> lowerBoundary )  

        /// Returns true if the reference does not lie in the boundary of the area.
        let referenceExceedsArea reference area = 
            referenceExceedsAreaRight reference area
            ||
            referenceExceedsAreaLeft reference area
            ||
            referenceExceedsAreaAbove reference area
            ||
            referenceExceedsAreaBelow reference area
 
        /// Returns true if the A1:A1-style area is of correct format.
        let isCorrect area = 
            try
                let hor = leftBoundary  area <= rightBoundary area
                let ver = upperBoundary area <= lowerBoundary area 

                if not hor then printfn "Right area boundary must be higher or equal to left area boundary."
                if not ver then printfn "Lower area boundary must be higher or equal to upper area boundary."

                hor && ver
                                        
            with
            | err -> 
                printfn "Area \"%s\" could not be parsed: %s" area.Value err.Message
                false

    /// Functions for working with the columns of a table.
    module TableColumns = 

        /// Gets the tableColumns from a table.
        let get (table : Table) =
            table.TableColumns

        /// Gets the columns from a tableColumns element.
        let getTableColumns (tableColumns : TableColumns) =
            tableColumns.Elements<TableColumn>()

        /// Retruns the number of columns in a tableColumns element.
        let count (tableColumns : TableColumns) =
            getTableColumns tableColumns
            |> Seq.length

        /// Creates a tableColumns sequence.
        let create (tableColumns : TableColumn seq) = 
            TableColumns(tableColumns |> Seq.map (fun c -> c :> OpenXmlElement))
            
    /// Functions for working with a single tableColumn of a table.
    module TableColumn =
        
        /// Gets Name of a tableColumn.
        let getName (tableColumn : TableColumn) =
            tableColumn.Name.Value

        /// Gets 1-based column index of tableColumn.
        let getId (tableColumn : TableColumn) =
            tableColumn.Id.Value

        /// Creates a tableColumn of the given name and id.
        let create (id : uint) (name : string) =
            let tc = TableColumn()
            tc.Name <- StringValue(name)
            tc.Id <- UInt32Value(id)
            tc

    /// Returns a table for which the predicate applied to its name returns true.
    let getByNameBy (predicate : string -> bool) (worksheetPart : WorksheetPart) =
        worksheetPart.TableDefinitionParts
        |> Seq.pick (fun t -> if predicate t.Table.Name.Value then Some t.Table else None)
    
    /// If a table exists, for which the predicate applied to its name returns true, gets it. Else returns None.
    let tryGetByNameBy (predicate : string -> bool) (worksheetPart : WorksheetPart) =
        worksheetPart.TableDefinitionParts
        |> Seq.tryPick (fun t -> if predicate t.Table.Name.Value then Some t.Table else None)

    /// If the worksheetPart contains tables, returns the first one. Else returns None.
    let tryGetFirst (worksheetPart : WorksheetPart) =
        try 
            worksheetPart.TableDefinitionParts
            |> Seq.head
            |> fun t ->  Some t.Table
        with
        | _ -> None

    /// List all tables contained in the worksheet.
    let list (worksheetPart : WorksheetPart) =
        worksheetPart.TableDefinitionParts
        |> Seq.map (fun t -> t.Table)

    /// Gets the name of a table.
    let getName (table : Table) =
        table.Name.Value

    /// Sets the name of a table.
    let setName (name : string) (table : Table) =
        table.Name <- (StringValue(name))
        table

    /// Gets the diplayName of a table.
    let getDisplayName (table:Table) =
        table.DisplayName.Value

    /// Sets the displayName of the table.
    let setDisplayName (displayName : string) (table : Table) =
        table.DisplayName <- (StringValue(displayName))
        table

    /// Gets the area of the table.
    let getArea (table : Table) =
        table.Reference

    /// Sets the area of the table.
    let setArea area (table : Table) =
        table.Reference <- area

    /// Creates a table from a name a area and tableColumns.
    let create (displayName : string) area tableColumns = 
        let t = Table()
        let a = AutoFilter()
        a.Reference <- area
        t.AutoFilter <- a
        t.Reference <- area
        t.DisplayName <- StringValue(displayName)
        t.TableColumns <- TableColumns.create tableColumns
        t

    /// Adds a table to the worksheetPart.
    let addTable (workbookPart : WorkbookPart) (worksheetPart : WorksheetPart) (table : Table) =
        // https://stackoverflow.com/questions/53440352/openxml-table-creation-how-do-i-create-tables-without-requiring-excel-to-repa
        let id = workbookPart.WorksheetParts |> Seq.sumBy (fun wsp -> wsp.TableDefinitionParts |> Seq.length) |> (+) 1
        let tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>(sprintf "rId%i" id);
        table.Id <- (UInt32Value.FromUInt32(uint id))
        table.Name <- StringValue(sprintf "Table%i" id)
        if table.TableStyleInfo = null then
            table.TableStyleInfo <- TableStyleInfo(Name = StringValue("TableStyleMedium7"), ShowFirstColumn = BooleanValue(false), ShowLastColumn = BooleanValue(false), ShowRowStripes = BooleanValue(true), ShowColumnStripes = BooleanValue(false) )
        tableDefinitionPart.Table <- table

        let tablePart = TablePart(Id = (StringValue (sprintf "rId%i" id)))
        let tableParts =             
            let tablePartSOption = 
                worksheetPart.Worksheet.ChildElements 
                |> Seq.tryPick (fun ce -> try ce :?> TableParts |> Some with | _ -> None) 

            match tablePartSOption with
            | Some tp -> tp
            | None ->
                let tableParts = new TableParts()
                tableParts.Count <- UInt32Value.FromUInt32 0u
                worksheetPart.Worksheet.AppendChild(tableParts) |> ignore
                tableParts

        tableParts.Count <- tableParts.Count.Value |> (+) 1u |> UInt32Value.FromUInt32     
        //tableParts.AppendChild(tablePart) |> ignore
        tableParts.AppendChild(tablePart) |> ignore

        //printfn "TableID: %i" id
        worksheetPart

    /// Create a table object by an area. If the first row of this area contains values in the given sheet, these are chosen as headers for the table and a table is returned.
    let tryCreateWithExistingHeaders (sst : SharedStringTable Option) sheetData name area =
        if Area.isCorrect area then
            try 
                let columns = 
                    let r = Area.upperBoundary area
                    [Area.leftBoundary area .. Area.rightBoundary area]
                    |> List.mapi (fun i c -> 
                        SheetData.getCellValueAt sst r c sheetData
                        |> TableColumn.create (i + 1 |> uint)                   
                    )
                create name area columns
                |> Some
            with
            | err -> 
                printfn "Could not retrieve table headers: %s" err.Message
                None
        else None


    /// Returns the headers of the columns.
    let getColumnHeaders (table : Table) = 
        table.TableColumns
        |> TableColumns.getTableColumns
        |> Seq.map (TableColumn.getName)

    /// Returns the tableColumn for which the predicate returns true if it exists. Else returns None.
    let tryGetTableColumnBy (predicate : TableColumn -> bool) (table : Table) =
        table.TableColumns
        |> TableColumns.getTableColumns
        |> Seq.tryFind predicate

    /// If a tableColumn with the given name exists in the table, returns it. Else returns None.
    let tryGetTableColumnByName name (table : Table) =
        table.TableColumns
        |> TableColumns.getTableColumns
        |> Seq.tryFind (TableColumn.getName >> (=) name)

    /// If a column with the given header exists in the table, returns its values. Else returns None.
    let tryGetColumnValuesByColumnHeader (sst : SharedStringTable Option) sheetData columnHeader (table : Table) =
        let area = getArea table
        table.TableColumns
        |> TableColumns.getTableColumns
        |> Seq.tryFindIndex (TableColumn.getName >> (=) columnHeader)
        |> Option.map (fun i ->
            let columnIndex =  (Area.leftBoundary area) + (uint i)
            [(Area.upperBoundary area + 1u) .. Area.lowerBoundary area]
            |> List.choose (fun r -> SheetData.tryGetCellValueAt sst r columnIndex sheetData)
        )
     
    /// If a column with the given header exists in the table, returns its indexed values. Else returns None.
    let tryGetIndexedColumnValuesByColumnHeader (sst : SharedStringTable Option) sheetData columnHeader (table : Table) =
        let area = getArea table
        table.TableColumns
        |> TableColumns.getTableColumns
        |> Seq.tryFindIndex (TableColumn.getName >> (=) columnHeader)
        |> Option.map (fun i ->
            let columnIndex =  (Area.leftBoundary area) + (uint i)
            [(Area.upperBoundary area + 1u) .. Area.lowerBoundary area]
            |> List.indexed
            |> List.choose (fun (i,r) -> 
                match SheetData.tryGetCellValueAt sst r columnIndex sheetData with
                | Some v -> Some (i,v)
                | None -> None
            )
        )

    /// If a key column and a value with the given header exist in the table, returns a tuple list of keys and values (else returns None). Missing values get replaced with the given default value.
    let tryGetKeyValuesByColumnHeaders (sst : SharedStringTable Option) sheetData keyColHeader valColHeader defaultValue (table : Table) =
        let area = getArea table
        let tableCols = 
            table.TableColumns
            |> TableColumns.getTableColumns
        match Seq.tryFindIndex (TableColumn.getName >> (=) keyColHeader) tableCols, Seq.tryFindIndex (TableColumn.getName >> (=) valColHeader) tableCols with
        | Some ki, Some vi ->
            let cki = (Area.leftBoundary area) + (uint ki)
            let cvi = (Area.leftBoundary area) + (uint vi)
            [(Area.upperBoundary area + 1u) .. Area.lowerBoundary area]
            |> List.map (fun r -> 
                let k = SheetData.getCellValueAt sst r cki sheetData
                let v = SheetData.tryGetCellValueAt sst r cvi sheetData |> Option.defaultValue defaultValue
                k,v
            )
            |> Some
        | _ -> None

    /// Reads a complete table. Values are stored sparsely in a dictionary, with the key being a row index and column header tuple.
    let toSparseValueMatrix (sst : SharedStringTable Option) sheetData (table : Table) =
        let area = getArea table
        let dictionary = System.Collections.Generic.Dictionary<int*string,string>()
        [Area.leftBoundary area .. Area.rightBoundary area]
        |> List.iter (fun c ->
            let upperBoundary = Area.upperBoundary area
            let lowerBoundary = Area.lowerBoundary area
            let header = SheetData.tryGetCellValueAt sst upperBoundary c sheetData |> Option.get
            List.init (lowerBoundary - upperBoundary |> int) (fun i ->
                let r = uint i + upperBoundary + 1u
                match SheetData.tryGetCellValueAt sst r c sheetData with
                | Some v -> dictionary.Add((i,header),v)
                | None -> ()                              
            )
            |> ignore
        )
        dictionary
        