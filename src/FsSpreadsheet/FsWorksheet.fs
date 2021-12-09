namespace FsSpreadsheet

// Type based on the type XLWorksheet used in ClosedXml
type FsWorksheet (name) =
    
    let mutable _name = name
    
    let mutable _rows : FsRow list = []

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

    member self.SortRows() = _rows <- _rows |> List.sortBy (fun r -> r.Index)