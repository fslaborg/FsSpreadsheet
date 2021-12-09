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

type FsAddress(address : string) =

    let mutable _address = address

    new (colI,rowI) = FsAddress(CellReference.ofIndices (uint32 colI) (uint32 rowI))

    member self.Address 
        with get() = _address
        and set(address) = _address <- address

    member self.OfIndices(colIndex,rowIndex) = _address <- CellReference.ofIndices colIndex rowIndex

    member self.ToIndices() = _address |> CellReference.toIndices |> fun (c,r) -> int c, int r

    member self.ColumnIndex 
        with get() = CellReference.toColIndex _address |> int
        and set(colI) = _address <- CellReference.setColIndex (uint32 colI) _address

    member self.RowIndex
        with get() = CellReference.toRowIndex _address |> int
        and set(rowI) = _address <- CellReference.setRowIndex (uint32 rowI) _address