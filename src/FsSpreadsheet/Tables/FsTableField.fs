namespace FsSpreadsheet

[<AllowNullLiteral>]
type FsTableField (name : string, index : int, column : FsRangeColumn, totalsRowLabel, totalsRowFunction) = 

    let mutable _totalsRowsFunction = totalsRowFunction
    let mutable _totalsRowLabel = totalsRowLabel
    let mutable _column = column
    let mutable _index = index
    let mutable _name = name

    new (name : string) = FsTableField(name,0,null,null,null)
