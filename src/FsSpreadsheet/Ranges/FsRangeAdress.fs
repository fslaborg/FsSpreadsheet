namespace rec FsSpreadsheet

//  Helper functions for working with "A1:A1"-style table areas.
/// The areas marks the area in which the table lies. 
module Range =

    /// Given A1-based top left start and bottom right end indices, returns a "A1:A1"-style area-
    let ofBoundaries fromCellReference toCellReference = 
        sprintf "%s:%s" fromCellReference toCellReference

    /// Given a "A1:A1"-style area, returns A1-based cell start and end cellReferences.
    let toBoundaries (area : string) = 
        area.Split ':'
        |> fun a -> a.[0], a.[1]

    /// Gets the right boundary of the area.
    let rightBoundary (area : string) = 
        toBoundaries area
        |> snd
        |> CellReference.toIndices
        |> fst

    /// Gets the left boundary of the area.
    let leftBoundary (area : string) = 
        toBoundaries area
        |> fst
        |> CellReference.toIndices
        |> fst

    /// Gets the Upper boundary of the area.
    let upperBoundary (area : string) = 
        toBoundaries area
        |> fst
        |> CellReference.toIndices
        |> snd

    /// Gets the lower boundary of the area.
    let lowerBoundary (area : string) = 
        toBoundaries area
        |> snd
        |> CellReference.toIndices
        |> snd

    /// Moves both start and end of the area by the given amount (positive amount moves area to right and vice versa).
    let moveHorizontal amount (area : string) =
        area
        |> toBoundaries
        |> fun (f,t) -> CellReference.moveHorizontal amount f, CellReference.moveHorizontal amount t
        ||> ofBoundaries

    /// Moves both start and end of the area by the given amount (positive amount moves area to right and vice versa).
    let moveVertical amount (area : string) =
        area
        |> toBoundaries
        |> fun (f,t) -> CellReference.moveHorizontal amount f, CellReference.moveHorizontal amount t
        ||> ofBoundaries

    /// Extends the right boundary of the area by the given amount (positive amount increases area to right and vice versa).
    let extendRight amount (area : string) =
        area
        |> toBoundaries
        |> fun (f,t) -> f, CellReference.moveHorizontal amount t
        ||> ofBoundaries

    /// Extends the left boundary of the area by the given amount (positive amount decreases the area to left and vice versa).
    let extendLeft amount (area : string) =
        area
        |> toBoundaries
        |> fun (f,t) -> CellReference.moveHorizontal amount f, t
        ||> ofBoundaries

    /// Returns true if the column index of the reference exceeds the right boundary of the area.
    let referenceExceedsAreaRight reference area = 
        (reference |> CellReference.toIndices |> fst) 
            > (area |> rightBoundary)
    
    /// Returns true if the column index of the reference exceeds the left boundary of the area.
    let referenceExceedsAreaLeft reference area = 
        (reference |> CellReference.toIndices |> fst) 
            < (area |> leftBoundary)  
 
    /// Returns true if the column index of the reference exceeds the upper boundary of the area.
    let referenceExceedsAreaAbove reference area = 
        (reference |> CellReference.toIndices |> snd) 
            > (area |> upperBoundary)
    
    /// Returns true if the column index of the reference exceeds the lower boundary of the area.
    let referenceExceedsAreaBelow reference area = 
        (reference |> CellReference.toIndices |> snd) 
            < (area |> lowerBoundary )  

    /// Returns true if the reference does not lie in the boundary of the area.
    let referenceExceedsArea reference area = 
        referenceExceedsAreaRight reference area
        ||
        referenceExceedsAreaLeft reference area
        ||
        referenceExceedsAreaAbove reference area
        ||
        referenceExceedsAreaBelow reference area
 
    /// Returns true if the A1:A1-style area is of correct format.
    let isCorrect area = 
        try
            let hor = leftBoundary  area <= rightBoundary area
            let ver = upperBoundary area <= lowerBoundary area 

            if not hor then printfn "Right area boundary must be higher or equal to left area boundary."
            if not ver then printfn "Lower area boundary must be higher or equal to upper area boundary."

            hor && ver
                                    
        with
        | err -> 
            printfn "Area \"%s\" could not be parsed: %s" area err.Message
            false

[<AllowNullLiteral>]
type FsRangeAddress(firstAddress : FsAddress, lastAddress : FsAddress) =

    let mutable _firstAddress = firstAddress
    let mutable _lastAddress = lastAddress


    new(rangeAddress) =
        let firstAdress,lastAddress = Range.toBoundaries rangeAddress
        FsRangeAddress(FsAddress(firstAdress),FsAddress(lastAddress))

    member self.Normalize () =

        let firstRow,lastRow = 
            if firstAddress.RowNumber < lastAddress.RowNumber then 
                firstAddress.RowNumber,lastAddress.RowNumber 
                else lastAddress.RowNumber,firstAddress.RowNumber

        let firstColumn,lastColumn = 
            if firstAddress.RowNumber < lastAddress.RowNumber then 
                firstAddress.RowNumber,lastAddress.RowNumber 
                else lastAddress.RowNumber,firstAddress.RowNumber

        _firstAddress <- FsAddress(firstRow,firstColumn)
        _lastAddress <- FsAddress(lastRow,lastColumn)

    member self.Range 
        with get() = Range.ofBoundaries _firstAddress.Address _lastAddress.Address
        and set(address) = 
            let firstAdress, lastAdress = Range.toBoundaries address            
            _firstAddress <- FsAddress (firstAdress)
            _lastAddress <- FsAddress (lastAdress)

    override self.ToString() =
        self.Range

    member self.FirstAddress = _firstAddress

    member self.LastAddress = _lastAddress