namespace FsSpreadsheet


/// Module containing functions to work with "A1" style excel cell references.
module CellReference =
    
    /// Transforms excel column string indices (e.g. A, B, Z, AA, CD) to index number (starting with A = 1).
    let colAdressToIndex (columnAdress : string) =
        let length = columnAdress.Length
        let mutable sum = 0u
        for i=0 to length-1 do
            let c = columnAdress.[length-1-i] |> System.Char.ToUpper
            let factor = 26. ** (float i) |> uint
            sum <- sum + ((uint c - 64u) * factor)
        sum

    /// Transforms number index to excel column string indices (e.g. A, B, Z, AA, CD) (starting with A = 1).
    let indexToColAdress i =
        let sb = System.Text.StringBuilder()
        let rec loop residual = 
            if residual = 0u then
                sb.ToString()
            else
                let modulo = (residual - 1u) % 26u
                sb.Insert(0, char (modulo + 65u)) |> ignore
                loop ((residual - modulo) / 26u)
        loop i
        

    /// Maps 1 based column and row indices to "A1" style reference.
    let ofIndices column (row : uint32) = 
        sprintf "%s%i" (indexToColAdress column) row

    /// Maps a "A1" style excel cell reference to a column * row index tuple (1 Based indices).
    let toIndices (reference : string) = 
        let inp = reference.ToUpper()
        let pattern = "([A-Z]*)(\d*)"
        let regex = System.Text.RegularExpressions.Regex.Match(inp,pattern)
        
        if regex.Success then
            regex.Groups
            |> fun a -> colAdressToIndex a.[1].Value, uint32 a.[2].Value
        else 
            failwithf "Reference %s does not match Excel A1-style" reference

    /// Maps a "A1" style excel cell reference to a column (1 Based indices).
    let toColIndex (reference : string) = 
        reference |> toIndices |> fst

    /// Maps a "A1" style excel cell reference to a row (1 Based indices).
    let toRowIndex (reference : string) = 
        reference |> toIndices |> snd

    /// Sets the column index (1 Based indices) of a "A1" style excel cell reference.
    let setColIndex colI (reference : string) = 
        reference |> toIndices |> fun (_,r) -> ofIndices colI r

    /// Sets the row index (1 Based indices) of a "A1" style excel cell reference.
    let setRowIndex rowI (reference : string) = 
        reference |> toIndices |> fun (c,_) -> ofIndices c rowI

    /// Changes the column portion of a "A1"-style reference by the given amount (positive amount = increase and vice versa).
    let moveHorizontal amount reference = 
        reference
        |> toIndices
        |> fun (c,r) -> (int64 c) + (int64 amount) |> uint32, r
        ||> ofIndices

    /// Changes the row portion of a "A1"-style reference by the given amount (positive amount = increase and vice versa).
    let moveVertical amount reference = 
        reference
        |> toIndices
        |> fun (c,r) -> c, (int64 r) + (int64 amount) |> uint32
        ||> ofIndices

type FsAddress(rowNumber : int, columnNumber : int, fixedRow : bool, fixedColumn : bool) =

    let mutable _fixedRow     = fixedRow
    let mutable _fixedColumn  = fixedColumn
    let mutable _rowNumber    = rowNumber
    let mutable _columnNumber = columnNumber

    let mutable _trimmedAddress = ""

    // ----------------------
    // ALTERNATE CONSTRUCTORS
    // ----------------------

    new (rowNumber : int, columnLetter : string, fixedRow : bool, fixedColumn : bool) =
        FsAddress(rowNumber,CellReference.colAdressToIndex columnLetter |> int,fixedRow,fixedColumn)

    new (rowNumber : int, columnNumber : int) =
        FsAddress(rowNumber,columnNumber,false,false)

    new (cellAddressString : string) =
        let colIndex,rowIndex = CellReference.toIndices cellAddressString
        FsAddress(int rowIndex,int colIndex)
    
    
    // ----------
    // PROPERTIES
    // ----------

    member self.ColumnNumber
        with get() = _columnNumber
        and set(colI) = _columnNumber <- colI

    member self.RowNumber
        with get() = _rowNumber
        and set(rowI) = _rowNumber <- rowI

    member self.Address 
        with get() = CellReference.ofIndices (uint32 _columnNumber) (uint32 _rowNumber)
        and set(address) = 
            let column,row = CellReference.toIndices address
            _rowNumber <- int row
            _columnNumber <- int column

    member self.FixedRow = false
    member self.FixedColumn = false


    // -------
    // METHODS
    // -------

    member this.LOL () = 1
    //let mutable _address = address

    /// <summary>Updates the row- and columnIndex respective to the given indices.</summary>
    member self.OfIndices(rowIndex,colIndex) = 
        _columnNumber <- colIndex
        _rowNumber <- rowIndex