namespace FsSpreadsheet.Net

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml

open FsSpreadsheet

/// Functions for working with rows (unmanaged: spans and cell references do not get automatically updated).
module Row =
    
    //  Helper functions for working with "1:1" style row spans
    /// Functions for working with spans. The spans mark the column wise area in which the row lies. 
    module Spans =

        /// Given 1 based column start and end indices, returns a "1:1" style spans.
        let fromBoundaries fromColumnIndex toColumnIndex = 
            sprintf "%i:%i" fromColumnIndex toColumnIndex
            |> StringValue.FromString
            |> List.singleton
            |> ListValue

        /// Given a "1:1" style spans, returns 1 based column start and end indices.
        let toBoundaries (spans:ListValue<StringValue>) = 
            spans.Items
            |> Seq.head
            |> fun x -> x.Value.Split ':'
            |> fun a -> uint32 a.[0],uint32 a.[1]

        //let toBoundaries (spans:ListValue<StringValue>) = 
        //    spans.Items
        //    |> Seq.head
        //    |> fun x -> System.Text.RegularExpressions.Regex.Matches(x.Value,@"\d*")
        //    |> fun a -> uint32 a.[0].Value,uint32 a.[2].Value

        /// Gets the right boundary of the spans.
        let rightBoundary (spans:ListValue<StringValue>) = 
            toBoundaries spans
            |> snd

        /// Gets the left boundary of the spans.
        let leftBoundary (spans:ListValue<StringValue>) = 
            toBoundaries spans
            |> fst

        /// Moves both start and end of the spans by the given amount (positive amount moves spans to right and vice versa).
        let moveHorizontal amount (spans:ListValue<StringValue>) =
            spans
            |> toBoundaries
            |> fun (f,t) -> amount + f, amount + t
            ||> fromBoundaries

        /// Extends the right boundary of the spans by the given amount (positive amount increases spans to right and vice versa).
        let extendRight amount (spans:ListValue<StringValue>) =
            spans
            |> toBoundaries
            |> fun (f,t) -> f, amount + t
            ||> fromBoundaries

        /// Extends the left boundary of the spans by the given amount (positive amount decreases the spans to left and vice versa).
        let extendLeft amount (spans:ListValue<StringValue>) =
            spans
            |> toBoundaries
            |> fun (f,t) -> f - amount, t
            ||> fromBoundaries

        /// Returns true if the column index of the reference exceeds the right boundary of the spans.
        let referenceExceedsSpansToRight reference spans = 
            (reference |> CellReference.toIndices |> fst) 
                > (spans |> rightBoundary)
        
        /// Returns true if the column index of the reference exceeds the left boundary of the spans.
        let referenceExceedsSpansToLeft reference spans = 
            (reference |> CellReference.toIndices |> fst) 
                < (spans |> leftBoundary)  
     
        /// Returns true if the column index of the reference does not lie in the boundary of the spans.
        let referenceExceedsSpans reference spans = 
            referenceExceedsSpansToRight reference spans
            ||
            referenceExceedsSpansToLeft reference spans

    /// Creates an empty Row.
    let empty () = Row()

    /// Returns a sequence of cells contained in the row.
    let toCellSeq (row : Row) : seq<Cell> = row.Descendants<Cell>() 

    /// Returns true if the row contains no cells.
    let isEmpty (row : Row) = toCellSeq row |> Seq.length |> (=) 0
    
    /// Iterates through all cells of a row with the given function f.
    let iterCells (f : Cell -> unit) (row : Row) = 
        row
        |> toCellSeq
        |> Seq.iter f 
        row

    /// Applies the function f to all cells of a row.
    let mapCells (f : Cell -> Cell) (row : Row) = 
        row
        |> toCellSeq
        |> Seq.iter (f >> ignore)
        row

    /// Returns the first cell in the row for which the predicate returns true.
    let findCell (predicate : Cell -> bool) (row : Row) =
        row
        |> toCellSeq
        |> Seq.find predicate

    /// Inserts a cell into the row before a reference cell.
    let insertCellBefore newCell refCell (row : Row) = 
        row.InsertBefore(newCell, refCell) |> ignore
        row

    /// Returns the rowIndex of the row.
    let getIndex (row : Row) = row.RowIndex.Value

    /// Sets the rowIndex of the row
    let setIndex index (row:Row) =
        row.RowIndex <- UInt32Value.FromUInt32 index
        row

    /// Returns true if the row contains a cell with the given columnIndex.
    let containsCellAt columnIndex (row : Row) =
        row
        |> toCellSeq
        |> Seq.exists (Cell.getReference >> CellReference.toIndices >> fst >> (=) columnIndex)

    /// Returns cell with the given columnIndex.
    let getCellAt columnIndex (row : Row) =
        row
        |> toCellSeq
        |> Seq.find (Cell.getReference >> CellReference.toIndices >> fst >> (=) columnIndex)

    /// Returns cell with the given columnIndex if it exists, else returns none.
    let tryGetCellAt columnIndex (row : Row) =
        row
        |> toCellSeq
        |> Seq.tryFind (Cell.getReference >> CellReference.toIndices >> fst >> (=) columnIndex)

    /// Returns cell matching or exceeding the given column index if it exists, else returns none.
    let tryGetCellAfter columnIndex (row : Row) =
        row
        |> toCellSeq
        |> Seq.tryFind (Cell.getReference >> CellReference.toIndices >> fst >> (<=) columnIndex)

    /// Returns the spans of the row.
    let getSpan (row : Row) = row.Spans

    /// Sets the spans of the row.
    let setSpan spans (row : Row) = 
        row.Spans <- spans
        row

    /// Extends the right boundary of the spans of the row by the given amount (positive amount increases spans to right and vice versa).
    let extendSpanRight amount row = 
        getSpan row
        |> Spans.extendRight amount
        |> fun s -> setSpan s row

    /// Extends the left boundary of the spans of the row by the given amount (positive amount decreases the spans to left and vice versa).
    let extendSpanLeft amount row = 
        getSpan row
        |> Spans.extendLeft amount
        |> fun s -> setSpan s row

    /// Append cell to the end of the row.
    let appendCell (cell : Cell) (row : Row) = 
        row.AppendChild(cell) |> ignore
        row

    /// Creates a row from the given rowIndex, columnSpans, and cells.
    let create index spans (cells : Cell seq) = 
        Row(childElements = (cells |> Seq.map (fun x -> x :> OpenXmlElement)))
        |> setIndex index
        |> setSpan spans
        |> fun r -> r.CloneNode(true) :?> Row

    /// Removes the cell at the given columnIndex from the row.
    let removeCellAt index (row : Row) =
        getCellAt index row
        |> row.RemoveChild
        |> ignore
        row

    /// Removes the cell at the given columnIndex from the row if present. Returns none if not.
    let tryRemoveCellAt index (row:Row) =
        tryGetCellAt index row
        |> Option.map (fun cell -> 
            row.RemoveChild(cell) |> ignore
            row)

    /// If the row contains a value at the given index, returns it. Returns none if not.
    let tryGetValueAt (sst : SST Option) index (row : Row) =
        row
        |> tryGetCellAt index
        |> Option.bind (Cell.tryGetValue sst)

    /// Matches the rowSpan to the cell references inside the row.
    let updateRowSpan (row : Row) : Row=
        let columnIndices =
            row
            |> toCellSeq
            |> Seq.map (Cell.getReference >> CellReference.toIndices >> fst)
        row
        |> setSpan (Spans.fromBoundaries (Seq.min columnIndices) (Seq.max columnIndices))

    /// Sets the rowIndex of the row and the row indices of the cells in the row to the given 1-based index.
    let updateRowIndex newIndex (row : Row) : Row =
        setIndex newIndex row
        |> mapCells (fun c -> 
            let (colI,rowI) = Cell.getReference c |> CellReference.toIndices
            Cell.setReference (CellReference.ofIndices colI newIndex) c
        )

    /// Creates a new row from the given values.
    let ofValues (sst : SharedStringTable Option) rowIndex (vals : 'T seq) =
        let spans = Spans.fromBoundaries 1u (Seq.length vals |> uint)
        vals
        |> Seq.mapi (fun i value -> 
            value
            |> Cell.fromValue sst (i + 1 |> uint) rowIndex
        )
        |> create rowIndex spans      

    /// If a cell with the given columnIndex exists in the row, moves it one column to the right.
    ///
    /// If there already was a cell at the new postion, moves that one too. Repeats until a value is moved into a position previously unoccupied.
    let rec moveValueBlockToRight columnIndex (row : Row) : Row=
        let span = getSpan row
        match row |> tryGetCellAt columnIndex  with
        | Some cell ->
            moveValueBlockToRight (columnIndex+1u) row |> ignore
            let newReference = (Cell.getReference cell |> CellReference.moveHorizontal 1)            
            Cell.setReference newReference cell |> ignore
            if span |> Spans.referenceExceedsSpansToRight newReference then
                extendSpanRight 1u row
            else row
        | None -> row

    /// Moves all cells starting with the given columnIndex in the row to the right by the given offset.
    let moveValuesToRight columnIndex (offset : uint32) (row : Row) : Row=
        row
        |> iterCells (fun cell -> 

            let cellCol,cellRow = 
                cell
                |> Cell.getReference
                |> CellReference.toIndices

            if cellCol >= columnIndex then
                cell
                |> Cell.setReference (CellReference.ofIndices (cellCol+offset) cellRow)
                |> ignore
        )
        |> extendSpanRight offset

    /// Maps the cells of the given row to tuples of 1-based column indices and the value strings using a sharedStringTable.
    let getIndexedValues (sst : SST Option) (row : Row) =
        row
        |> toCellSeq
        |> Seq.choose (fun cell -> 
            Cell.tryGetValue sst cell
            |> Option.map (fun v -> 
                cell 
                |> Cell.getReference 
                |> CellReference.toIndices 
                |> fst,
                v
            )            
        )


    /// Maps the cells of the given row to the value strings.
    let getRowValues (sst : SST Option) (row : Row)  =
        row
        |> toCellSeq
        |> Seq.choose (Cell.tryGetValue sst)

    /// Maps each cell of the given row to each respective value strings if it exists, else returns None.
    let tryGetRowValues (sst : SST option) (row : Row) =
        toCellSeq row
        |> Seq.map (Cell.tryGetValue sst)

    /// Maps the cells of the given row to the value strings for all existing cells.
    let getPresentRowValues (sst : SST option) (row : Row) =
        toCellSeq row
        |> Seq.choose (Cell.tryGetValue sst)

    
    /// Adds a value as a cell to the row at the given columnIndex.
    ///
    /// If a cell exists at the given columnIndex, shoves it to the right.
    let insertValue (sst : SharedStringTable Option) index (value : 'T) (row : Row) = 

        let refCell = row |> tryGetCellAfter index 
        let cell = Cell.fromValue sst index (getIndex row) value

        match refCell with
        | Some ref -> 
            row
            |> moveValuesToRight index 1u 
            |> insertCellBefore cell ref
        | None ->
            let spans = getSpan row
            let spanExceedance = index - (spans |> Spans.rightBoundary)
                           
            row
            |> extendSpanRight spanExceedance
            |> appendCell cell


    /// Adds a value as a cell to the row at the given columnindex using a sharedStringTable.
    ///
    /// If a cell exists at the given columnindex, shoves it to the right.
    let insertValueAt (sst : SharedStringTable Option) index (value : 'T) (row : Row) = 

        let refCell = row |> tryGetCellAfter index 
        let cell = Cell.fromValue sst index (getIndex row) value

        match refCell with
        | Some ref -> 
            row
            |> moveValuesToRight index 1u
            |> insertCellBefore cell ref
        | None ->
            let spans = getSpan row
            let spanExceedance = index - (spans |> Spans.rightBoundary)
                               
            row
            |> extendSpanRight spanExceedance
            |> appendCell cell

    
    /// Adds a value as a cell to the end of the row.
    let appendValue (sst : SharedStringTable Option) (value : 'T) (row : Row) = 
        let colIndex = 
            row
            |> getSpan
            |> Spans.rightBoundary
        let cell = Cell.fromValue sst (colIndex + 1u) (row |> getIndex) value
        row
        |> appendCell cell
        |> extendSpanRight 1u 

    
    /// Add a value as a cell to the row at the given columnIndex.
    ///
    /// If a cell exists at the given columnIndex, overwrites it.
    // To-Do: Add version using a sharedStringTable
    let setValue (sst : SharedStringTable Option) index (value : 'T) (row : Row) = 

        let refCell = row |> tryGetCellAfter index 
        let cell = Cell.fromValue sst index (getIndex row) value

        match refCell with
        | Some ref when Cell.getReference ref = Cell.getReference cell ->
            ref  |> Cell.setType (Cell.getType cell)  |> ignore
            ref |> Cell.setValue ((Cell.getCellValue cell).Clone() :?> CellValue) |> ignore
            row 
        | Some ref -> 
            row |> insertCellBefore cell ref
        | None ->
            let spans = getSpan row
            let spanExceedance = index - (spans |> Spans.rightBoundary)
            row
            |> extendSpanRight spanExceedance
            |> appendCell cell

    /// Includes a value from a sharedStringTable in the cells of the row.
    let includeSharedStringValue (sst : SST) (row : Row) =
        row
        |> mapCells (Cell.includeSharedStringValue sst)
