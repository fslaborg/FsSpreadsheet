namespace FsSpreadsheet

[<AllowNullLiteral>]
type FsRangeColumn(rangeAddress) =

    inherit FsRangeBase(rangeAddress)
    
    //new () = 
    //    let range = FsRangeAddress(FsAddress(0,0),FsAddress(0,0))
    //    FsRangeColumn(range)
    
    new (index) = FsRangeColumn (FsRangeAddress(FsAddress(0,index),FsAddress(0,index)))

    member self.Cell(rowIndex,cells) = base.Cell(FsAddress(rowIndex - base.RangeAddress.FirstAddress.RowNumber + 1,1),cells)
    
    member self.FirstCell(cells : FsCellsCollection) = 
        //let firstAddrRow, firstAddrCol = base.RangeAddress.FirstAddress |> fun fa -> fa.RowNumber, fa.ColumnNumber
        //base.Cell(FsAddress(firstAddrRow, firstAddrCol), cells)
        base.Cell(FsAddress(1, 1), cells)

    member self.Cells(cells) = base.Cells(cells)

    member self.Index 
        with get() = self.RangeAddress.FirstAddress.ColumnNumber
        and set(i) = 
            self.RangeAddress.FirstAddress.ColumnNumber <- i
            self.RangeAddress.LastAddress.ColumnNumber <- i
