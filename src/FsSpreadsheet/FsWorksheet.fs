namespace FsSpreadsheet

// Type based on the type XLWorksheet used in ClosedXml
type FsWorksheet (name) =
    
    let mutable _name = name
    
    let mutable _rows : FsRow list = []

    let mutable _tables : FsTable list = []

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

    member self.Table(tableName,rangeAddress) = 
        match _tables |> List.tryFind (fun table -> table.Name = name) with
        | Some table ->
            table
        | None -> 
            let table = FsTable(tableName,rangeAddress)
            _tables <- List.append _tables [table]
            table
        
    member self.GetTables() = _tables

    member self.SortRows() = _rows <- _rows |> List.sortBy (fun r -> r.Index)