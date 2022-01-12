namespace FsSpreadsheet

// Type based on the type XLWorksheet used in ClosedXml
type FsWorksheet (name) =
    
    let mutable _name = name
    
    let mutable _rows : FsRow list = []

    let mutable _tables : FsTable list = []

    let mutable _cells = FsCellsCollection()

    new () = FsWorksheet("")

    member self.Name
        with get() = _name
        and set(name) = _name <- name

    member self.Row(rowIndex) = 
        match _rows |> List.tryFind (fun row -> row.Index = rowIndex) with
        | Some row ->
            row
        | None -> 
            let row = FsRow(rowIndex) 
            _rows <- List.append _rows [row]
            row
        
    member self.GetRows() = _rows

    member self.Table(tableName,rangeAddress,showHeaderRow) = 
        match _tables |> List.tryFind (fun table -> table.Name = name) with
        | Some table ->
            table
        | None -> 
            let table = FsTable(tableName,rangeAddress,showHeaderRow)
            _tables <- List.append _tables [table]
            table
      
    member self.Table(tableName,rangeAddress) = 
        self.Table(tableName,rangeAddress,true)

    member self.RescanRows() =
        ()

    member self.CellCollection
        with get () = _cells

    member self.GetTables() = _tables

    member self.SortRows() = _rows <- _rows |> List.sortBy (fun r -> r.Index)