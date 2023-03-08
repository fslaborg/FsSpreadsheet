namespace FsSpreadsheet

// Type based on the type XLRow used in ClosedXml
/// <summary>Creates an FsRow from the given FsRangeAddress, consisting of FsCells within a given FsCellsCollection, and a styleValue.</summary>
/// <remarks>The FsCellsCollection must only cover 1 row!</remarks>
/// <exception cref="System.Exception">if given FsCellsCollection has more than 1 row.</exception>
type FsRow (rangeAddress : FsRangeAddress, cells : FsCellsCollection, styleValue)= 

    inherit FsRangeBase(rangeAddress,styleValue)

    let cells = 
        let rowCount =
            cells.GetCells()
            |> Seq.map (fun c -> c.Address.RowNumber)
            |> Seq.distinct
            |> Seq.length
        if rowCount > 1 then
            failwith "FsCellsCollection must only have 1 row to be qualified for FsRow creation."
        cells

    new () = FsRow (FsRangeAddress(FsAddress(0,0),FsAddress(0,0)),FsCellsCollection(),null)

    // ask HLW: isn't this wrong? The FsRangeAddress will reach from col 1 to col 1 but shouldn't it be reaching from col 1 to how long the FsCellsCollection is?
    // or even from first colIndex in the FsCellsCollection to last colIndex of it?
    new (index,cells) = FsRow (FsRangeAddress(FsAddress(index,1),FsAddress(index,1)),cells,null)


    // ----------
    // PROPERTIES
    // ----------

    /// The associated FsCells.
    member self.Cells = 
        base.Cells(cells)

    /// The index of the FsRow.
    member self.Index 
        with get() = self.RangeAddress.FirstAddress.RowNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.RowNumber <- i
            self.RangeAddress.LastAddress.RowNumber <- i


    // -------
    // METHODS
    // -------

    /// Returns the index of the given FsRow.
    static member getIndex (row : FsRow) = 
        row.Index

    /// Returns the FsCell at columnIndex.
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

    /// <summary>Returns the FsCell at the given columnIndex from an FsRow.
    static member getCellAt colIndex (row : FsRow) =
        row.Cell(colIndex)

    /// Inserts the value at columnIndex as an FsCell. If there is an FsCell at the position, this FsCells and all the ones right to it are shifted to the right.
    member this.InsertValueAt(colIndex, (value : 'a)) =
        let cell = FsCell(value)
        cells.Add(int32 this.Index, int32 colIndex, cell)

    /// Adds a value at the given row- and columnIndex to FsRow using.
    ///
    /// If a cell exists in the given position, shoves it to the right.
    static member insertValueAt colIndex value (row : FsRow) =
        row.InsertValueAt(colIndex, value)
        row

    //member self.SortCells() = _cells <- _cells |> List.sortBy (fun c -> c.WorksheetColumn)