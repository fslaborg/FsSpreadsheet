namespace FsSpreadsheet


type FsRange(rangeAddress : FsRangeAddress, styleValue) = 

    inherit FsRangeBase(rangeAddress)

    member this.A = 1

