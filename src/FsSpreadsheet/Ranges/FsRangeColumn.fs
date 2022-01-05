namespace FsSpreadsheet

[<AllowNullLiteral>]
type FsRangeColumn(rangeAddress) =

    inherit FsRangeBase(rangeAddress)
    
    //member self.Cell