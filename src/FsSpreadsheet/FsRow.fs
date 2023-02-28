namespace FsSpreadsheet

// Type based on the type XLRow used in ClosedXml
type FsRow (rangeAddress : FsRangeAddress, cells : FsCellsCollection, styleValue)= 

    inherit FsRangeBase(rangeAddress,styleValue)

    let cells = cells

    new () = FsRow (FsRangeAddress(FsAddress(0,0),FsAddress(0,0)),FsCellsCollection(),null)

    new (index,cells) = FsRow (FsRangeAddress(FsAddress(index,1),FsAddress(index,1)),cells,null)

    member self.Cell(columnIndex) = base.Cell(FsAddress(1,columnIndex),cells)
        
        //match _cells |> List.tryFind (fun cell -> cell.WorksheetColumn = columnIndex) with
        //| Some cell ->
        //    cell
        //| None -> 
        //    let cell = FsCell()
        //    cell.WorksheetColumn <- columnIndex
        //    cell.WorksheetRow <- _index
        //    _cells <- List.append _cells [cell]
        //    cell

    member self.Cells = base.Cells(cells)

    member self.Index 
        with get() = self.RangeAddress.FirstAddress.RowNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.RowNumber <- i
            self.RangeAddress.LastAddress.RowNumber <- i

    //member self.SortCells() = _cells <- _cells |> List.sortBy (fun c -> c.WorksheetColumn)

    static member getIndex (row : FsRow) = 
        row.Index