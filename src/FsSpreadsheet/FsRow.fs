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

    member this.InsertValueAt(colIndex, (value : 'a)) =
        let cell = FsCell(value)
        cells.Add(int32 this.Index, int32 colIndex, cell)

    member self.Cells = base.Cells(cells)

    /// The index of the FsRow.
    member self.Index 
        with get() = self.RangeAddress.FirstAddress.RowNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.RowNumber <- i
            self.RangeAddress.LastAddress.RowNumber <- i

    //member self.SortCells() = _cells <- _cells |> List.sortBy (fun c -> c.WorksheetColumn)

    /// Returns the index of the given FsRow.
    static member getIndex (row : FsRow) = 
        row.Index

    /// Add a value at the given row- and columnindex to FsRow using.
    ///
    /// If a cell exists in the given position, shoves it to the right.
    static member insertValueAt colIndex value (row : FsRow) =
        row.InsertValueAt(colIndex, value)
        row