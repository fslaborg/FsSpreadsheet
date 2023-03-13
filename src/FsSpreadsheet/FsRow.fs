namespace FsSpreadsheet

// Type based on the type XLRow used in ClosedXml
/// <summary>Creates an FsRow from the given FsRangeAddress, consisting of FsCells within a given FsCellsCollection, and a styleValue.</summary>
/// <remarks>The FsCellsCollection must only cover 1 row!</remarks>
/// <exception cref="System.Exception">if given FsCellsCollection has more than 1 row.</exception>
type FsRow (rangeAddress : FsRangeAddress, cells : FsCellsCollection, styleValue)= 

    inherit FsRangeBase(rangeAddress,styleValue)

    let cells = cells

    new () = FsRow (FsRangeAddress(FsAddress(0,0),FsAddress(0,0)),FsCellsCollection(),null)

    /// <summary>Create an FsRow from a given FsCellsCollection and an rowIndex.</summary>
    /// <remarks>The appropriate range of the cells (i.e. minimum colIndex and maximum colIndex) is derived from the FsCells with the matching rowIndex.</remarks>
    new (index, (cells : FsCellsCollection)) = 
        let minColIndex = (cells.GetCellsInRow index |> Seq.minBy (fun c -> c.Address.ColumnNumber)).Address.ColumnNumber
        let maxColIndex = (cells.GetCellsInRow index |> Seq.maxBy (fun c -> c.Address.ColumnNumber)).Address.ColumnNumber
        FsRow (FsRangeAddress(FsAddress(index, minColIndex),FsAddress(index, maxColIndex)), cells, null)


    // ----------
    // PROPERTIES
    // ----------

    /// The associated FsCells.
    member self.Cells = 
        base.Cells(cells)

    /// <summary>The index of the FsRow.</summary>
    member self.Index 
        with get() = self.RangeAddress.FirstAddress.RowNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.RowNumber <- i
            self.RangeAddress.LastAddress.RowNumber <- i


    // -------
    // METHODS
    // -------

    /// <summary>Returns the index of the given FsRow.</summary>
    static member getIndex (row : FsRow) = 
        row.Index

    /// <summary>Returns the FsCell at columnIndex.</summary>
    member self.Cell(columnIndex) = 
        base.Cell(FsAddress(1,columnIndex),cells)
        
        //match _cells |> List.tryFind (fun cell -> cell.WorksheetColumn = columnIndex) with
        //| Some cell ->
        //    cell
        //| None -> 
        //    let cell = FsCell()
        //    cell.WorksheetColumn <- columnIndex
        //    cell.WorksheetRow <- _index
        //    _cells <- List.append _cells [cell]
        //    cell

    /// <summary>Returns the FsCell at the given columnIndex from an FsRow.</summary>
    static member getCellAt colIndex (row : FsRow) =
        row.Cell(colIndex)

    /// Inserts the value at columnIndex as an FsCell. If there is an FsCell at the position, this FsCells and all the ones right to it are shifted to the right.
    member this.InsertValueAt(colIndex, (value : 'a)) =
        let cell = FsCell(value)
        cells.Add(int32 this.Index, int32 colIndex, cell)

    /// <summary>Adds a value at the given row- and columnIndex to FsRow using.
    ///
    /// If a cell exists in the given position, shoves it to the right.</summary>
    static member insertValueAt colIndex value (row : FsRow) =
        row.InsertValueAt(colIndex, value) |> ignore
        row

    //member self.SortCells() = _cells <- _cells |> List.sortBy (fun c -> c.WorksheetColumn)

    // TO DO (later)
    ///// Takes an FsCellsCollection and creates an FsRow from the given rowIndex and the cells in the FsCellsCollection that share the same rowIndex.
    //static member fromCellsCollection rowIndex (cellsCollection : FsCellsCollection) =