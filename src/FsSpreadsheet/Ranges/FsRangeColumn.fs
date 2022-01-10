namespace FsSpreadsheet

[<AllowNullLiteral>]
type FsRangeColumn(rangeAddress) =

    inherit FsRangeBase(rangeAddress)
    
    //new () = 
    //    let range = FsRangeAddress(FsAddress(0,0),FsAddress(0,0))
    //    FsRangeColumn(range)
    
    new (index) = FsRangeColumn (FsRangeAddress(FsAddress(0,index),FsAddress(0,index)))

    member self.Cell(columnIndex,cells) = base.Cell(FsAddress(1,columnIndex),cells)
    
    member self.FirstCell(cells : FsCellsCollection) = 
        cells.TryGetCell(base.RangeAddress.FirstAddress.RowNumber,base.RangeAddress.LastAddress.ColumnNumber)
        |> Option.get

    member self.Cells(cells) = base.Cells(cells)

    member self.Index 
        with get() = self.RangeAddress.FirstAddress.ColumnNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.ColumnNumber <- i
            self.RangeAddress.LastAddress.ColumnNumber <- i
