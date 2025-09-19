namespace FsSpreadsheet

open Fable.Core

[<AllowNullLiteral>][<AttachMembers>]
type FsTableRow (rangeAddress : FsRangeAddress) = 

    inherit FsRangeRow(rangeAddress)