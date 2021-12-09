namespace FsSpreadsheet

// Type based on the type XLRow used in ClosedXml
type FsRow (index : int) = 

    let mutable _index = index
    let mutable _cells : FsCell list = []

    new () = FsRow (0)

    member self.Cell(columnIndex) = 
        match _cells |> List.tryFind (fun cell -> cell.WorksheetColumn = columnIndex) with
        | Some cell ->
            cell
        | None -> 
            let cell = FsCell()
            cell.WorksheetColumn <- columnIndex
            cell.WorksheetRow <- _index
            _cells <- List.append _cells [cell]
            cell

    member self.GetCells() = _cells

    member self.Index 
        with get() = _index
        and set(i) = _index <- i

    member self.SortCells() = _cells <- _cells |> List.sortBy (fun c -> c.WorksheetColumn)