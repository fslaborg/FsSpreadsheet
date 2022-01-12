namespace FsSpreadsheet

// Type based on the type XLRow used in ClosedXml
type FsRow (rangeAddress : FsRangeAddress, styleValue)= 

    inherit FsRangeBase(rangeAddress,styleValue)

    new () = FsRow (FsRangeAddress(FsAddress(0,0),FsAddress(0,0)),null)

    new (index) = FsRow (FsRangeAddress(FsAddress(index,1),FsAddress(index,1)),null)

    member self.Cell(columnIndex,cells) = base.Cell(FsAddress(1,columnIndex),cells)
        
        //match _cells |> List.tryFind (fun cell -> cell.WorksheetColumn = columnIndex) with
        //| Some cell ->
        //    cell
        //| None -> 
        //    let cell = FsCell()
        //    cell.WorksheetColumn <- columnIndex
        //    cell.WorksheetRow <- _index
        //    _cells <- List.append _cells [cell]
        //    cell

    member self.Cells(cells) = base.Cells(cells)

    member self.Index 
        with get() = self.RangeAddress.FirstAddress.RowNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.RowNumber <- i
            self.RangeAddress.LastAddress.RowNumber <- i

    //member self.SortCells() = _cells <- _cells |> List.sortBy (fun c -> c.WorksheetColumn)