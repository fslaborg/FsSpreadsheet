namespace FsSpreadsheet

[<AllowNullLiteral>]
type FsTableField (name : string, index : int, column : FsRangeColumn, totalsRowLabel, totalsRowFunction) = 

    let mutable _totalsRowsFunction = totalsRowFunction
    let mutable _totalsRowLabel = totalsRowLabel
    let mutable _column = column
    let mutable _index = index
    let mutable _name = name

    new (name : string) = FsTableField(name,0,null,null,null)

    new (name : string, index : int) = FsTableField(name,0,null,null,null)

    member this.Index 
        with get () = _index
        and set (index) = 
            if index = _index then
                ()
            else 
                _index <- index
                _column <- null

    member this.Column 
        with get () = 
            //let column =
            //    if _column = null then

            //    else               
                    _column
            //column
        and set (column) = _column <- column

    member this.Name = _name

    member this.SetName (name, cells : FsCellsCollection, showHeaderRow : bool) =
        _name <- name
        if showHeaderRow then
            this.Column.FirstCell(cells).SetValueAs<string>(name)
            |> ignore

    /// <summary>Returns the header cell for the table field.</summary>
    member this.HeaderCell (cells : FsCellsCollection, showHeaderRow : bool) =
        if not showHeaderRow then 
            failwithf "tried to get header cell of table field \"%s\" even though showHeaderRow is set to zero" _name
        else
            this.Column.FirstCell(cells)

    /// <summary>Gets the collection of data cells for this field. Excludes the header and footer cells.</summary>
    member this.DataCells (cells : FsCellsCollection, showHeaderRow : bool) =
        let predicate cell = not (showHeaderRow && this.HeaderCell(cells, showHeaderRow) <> cell)
        this.Column.Cells(cells, predicate)
