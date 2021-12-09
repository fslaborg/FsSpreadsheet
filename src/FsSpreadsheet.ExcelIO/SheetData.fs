namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml.Spreadsheet

open FsSpreadsheet

/// Functions for working with SheetData. (Unmanaged: Row indices and cell references do not automatically get updated)
module SheetData = 

    /// Creates an empty SheetData.
    let empty () = new SheetData()

    /// Inserts a row into the SheetData before a reference row.
    let insertBefore (row : Row) (refRow : Row) (sheetData : SheetData) = 
        sheetData.InsertBefore(row,refRow) |> ignore
        sheetData

    /// Append a row to the end of the SheetData.
    let appendRow (row : Row) (sheetData : SheetData) = 
        sheetData.AppendChild row |> ignore
        sheetData

    //----------------------------------------------------------------------------------------------------------------------
    //                                              Row(s)                                                                  
    //----------------------------------------------------------------------------------------------------------------------

    /// Returns a sequence of rows contained in the SheetData.
    let getRows (sheetData : SheetData) : seq<Row>= 
        sheetData.Descendants<Row>()

    /// Applies the given function f onto every row in the SheetData.
    let mapRows (f: Row -> Row) (sheetData : SheetData) =
        sheetData
        |> getRows
        |> Seq.iter (f >> ignore)
        sheetData


    /// Returns the number of rows contained in the SheetData.
    let countRows (sheetData : SheetData) = 
        getRows sheetData
        |> Seq.length
    
    /// Returns row matching or exceeding the given row index if it exists, else returns None.
    let tryGetRowAfter index (sheetData : SheetData) = 
        getRows sheetData
        |> Seq.tryFind (Row.getIndex >> (<=) index)

    /// Returns the row with the given rowIndex if it exists, else returns None.
    let tryGetRowAt index (sheetData : SheetData) = 
        getRows sheetData
        |> Seq.tryFind (Row.getIndex >> (=) index)

    /// Returns the row with the given rowIndex.
    let getRowAt index (sheetData : SheetData) = 
        getRows sheetData
        |> Seq.find (Row.getIndex >> (=) index)

    /// Returns true if the SheetData contains a row with the given rowIndex.
    let containsRowAt index (sheetData : SheetData) = 
        getRows sheetData
        |> Seq.exists (Row.getIndex >> (=) index)

    /// Removes the row at the given rowIndex.
    let removeRowAt index (sheetData : SheetData) = 
        let r = sheetData |> getRowAt index
        sheetData.RemoveChild(r) |> ignore
        sheetData

    /// Removes the row at the given rowIndex if it exists, else it returns None.
    let tryRemoveRowAt index (sheetData : SheetData) = 
        sheetData |> tryGetRowAt index
        |> Option.map (fun r ->
            sheetData.RemoveChild(r) |> ignore
            sheetData
        )

    ///If the row with index rowIndex exists in the sheet, moves it downwards by amount. Negative amounts will move the row upwards.
    let moveRowVertical (amount : int) rowIndex (sheetData : SheetData) = 
        match sheetData |> tryGetRowAt rowIndex with
        | Some row -> 
            let shift = (int64 (Row.getIndex row)) + (int64 amount)
            row
            |> Row.setIndex (uint32 shift)
            |> Row.toCellSeq
            |> Seq.iter (fun cell -> 
                cell
                |> Cell.setReference (
                    Cell.getReference cell 
                    |> CellReference.moveVertical amount
                ) 
                |> ignore            
            )
            sheetData
        | None -> 
            printfn "Warning: Row with index %i does not exist" rowIndex
            sheetData
    
    ///If a row with the given rowIndex exists in the sheet, moves it one position downwards. 
    ///
    /// If there already was a row at the new postion, moves that one too. Repeats until a row is moved into a position previously unoccupied.
    let rec moveRowBlockDownward rowIndex (sheetData : SheetData) =
        if sheetData |> containsRowAt rowIndex then             
            sheetData
            |> moveRowBlockDownward (rowIndex+1u)
            |> moveRowVertical 1 rowIndex
        else
            sheetData

    /// Returns the index of the last row in the sheet.
    let getMaxRowIndex (sheetData : SheetData) =
        getRows sheetData
        |> fun s -> 
            if Seq.isEmpty s then 
                0u
            else 
                s
                |> Seq.map (Row.getIndex)
                |> Seq.max

    
    /// Gets the string value of the cell at the given 1-based column and rowIndex, if it exists, else returns None.
    let tryGetRowValuesAt (sst : SharedStringTable Option) rowIndex (sheet : SheetData) =
        sheet 
        |> tryGetRowAt rowIndex
        |> Option.map (
            Row.toCellSeq
            >> Seq.map (Cell.getValue sst)
        )

    /// Gets the string values of the row at the given 1-based rowIndex.
    let getRowValuesAt (sst : SharedStringTable Option) rowIndex (sheet : SheetData) =
        sheet
        |> getRowAt rowIndex
        |> Row.toCellSeq
        |> Seq.map (Cell.getValue sst)

    /// Maps the cells of the given row to tuples of 1-based column indices and the value strings using a sharedStringTable, if it exists, else returns None.
    let tryGetIndexedRowValuesAt (sst : SharedStringTable Option) rowIndex (sheet : SheetData) =
        sheet
        |> tryGetRowAt rowIndex
        |> Option.map (
            Row.toCellSeq
            >> Seq.map (fun cell -> 
                cell
                |> Cell.getReference  
                |> CellReference.toIndices 
                |> fst,

                cell |> Cell.getValue sst)
        )

    
    /// Adds values as a row to the sheet at the given rowIndex with the given horizontal offset.
    ///
    /// If a row exists at the given rowIndex, shoves it downwards.
    let insertRowWithHorizontalOffsetAt (sst : SharedStringTable Option) (offset : int) (vals : 'T seq) rowIndex (sheet : SheetData) =
        let uiO = uint32 offset
        let spans = Row.Spans.fromBoundaries (uiO + 1u) (Seq.length vals |> uint32 |> (+) uiO )
        let newRow = 
            vals
            |> Seq.mapi (fun i v -> 
                Cell.fromValue sst ((int64 i) + 1L + (int64 offset) |> uint32) rowIndex v
            )
            |> Row.create (uint32 rowIndex) spans
        let refRow = tryGetRowAfter (uint rowIndex) sheet
        match refRow with
        | Some ref -> 
            sheet
            |> moveRowBlockDownward rowIndex 
            |> insertBefore newRow ref
        | None ->
            appendRow newRow sheet


  
    /// Adds values as a row to the sheet at the given rowIndex.
    ///
    /// If a row exists at the given rowIndex, shoves it downwards.
    let insertRowValuesAt (sst : SharedStringTable Option) (vals : 'T seq) rowIndex (sheet : SheetData) =
        insertRowWithHorizontalOffsetAt sst 0 vals rowIndex sheet


    /// Append the values as a row to the end of the sheet.
    let appendRowValues (sst : SharedStringTable Option) (vals : 'T seq) (sheet : SheetData) =
        let i = (getMaxRowIndex sheet) + 1u
        insertRowValuesAt sst vals i sheet
  

    /// Append the value as a cell to the end of the row.
    // To-Do: Add version using a sharedStringTable
    let appendValueToRowAt (sst : SharedStringTable Option) rowIndex (value : 'T) (sheet : SheetData) =
        match tryGetRowAt rowIndex sheet with
        | Some row -> 
            row
            |> Row.appendValue sst value
            |> ignore
            sheet
        | None -> insertRowValuesAt sst [value] rowIndex sheet    

    /// Removes row from sheet and move the following rows up.
    // To-Do: Add version using a sharedStringTable
    let deleteRowAt rowIndex (sheet : SheetData) : SheetData =
        sheet 
        |> removeRowAt rowIndex
        |> getRows
        |> Seq.filter (Row.getIndex >> (<) rowIndex)
        |> Seq.fold (fun sheetData row -> 
            moveRowVertical -1 (Row.getIndex row) sheetData           
        ) sheet


//----------------------------------------------------------------------------------------------------------------------
//                                              Cell(s)                                                                 
//----------------------------------------------------------------------------------------------------------------------

    /// Return a cell at the given row- and columnIndex in the SheetData if it exists. Else returns None.
    let tryGetCellAt (rowIndex : uint32) (columnIndex : uint32) (sheetData : SheetData) =         
        sheetData
        |> tryGetRowAt rowIndex 
        |> Option.bind (Row.tryGetCellAt columnIndex)

    /// Returns a cell at the given row- and columnIndex in the SheetData.
    let getCellAt (rowIndex: uint32) (columnIndex : uint32) (sheetData:SheetData) = 
        sheetData
        |> getRowAt rowIndex 
        |> Row.getCellAt columnIndex

    /// Gets the string value of the cell at the given 1-based column- and rowIndex using a sharedStringTable.
    let getCellValueAt (sst : SharedStringTable Option) (rowIndex : uint32) (columnIndex : uint32) (sheetData : SheetData) = 
        sheetData 
        |> getCellAt rowIndex columnIndex
        |> Cell.getValue sst

    /// Gets the string value of the cell at the given 1-based column- and rowIndex using a sharedStringTable if it exists. Else returns None.
    let tryGetCellValueAt (sst : SharedStringTable Option) (rowIndex: uint32) (columnIndex : uint32) (sheetData:SheetData) = 
        sheetData 
        |> tryGetCellAt rowIndex columnIndex
        |> Option.bind (Cell.tryGetValue sst)

    /// Add a value at the given row- and columnindex to sheet using a shared string table.
    ///
    /// If a cell exists in the given postion, shoves it to the right
    let insertValueAt (sst : SharedStringTable Option) rowIndex columnIndex (value : 'T) (sheet : SheetData) =
        match tryGetRowAt rowIndex sheet with
        | Some row -> 
            row
            |> Row.insertValue sst columnIndex value  
            |> ignore
            sheet
        | None -> insertRowWithHorizontalOffsetAt sst (columnIndex - 1u |> int) [value] rowIndex sheet


    /// Add a value at the given row- and columnIndex using a sharedStringTable.
    ///
    /// If a cell exists in the given postion, overwrites it.
    // To-Do: Add version using a sharedStringTable
    let setValueAt (sst : SharedStringTable Option) rowIndex columnIndex (value : 'T) (sheet : SheetData) =
        match tryGetRowAt rowIndex sheet with
        | Some row -> 
            row 
            |> Row.setValue sst columnIndex value 
            |> ignore
            sheet
        | None -> insertRowWithHorizontalOffsetAt sst (columnIndex - 1u |> int) [value] rowIndex sheet

    /// Removes the value at the given row- and columnIndex from the sheet if it exists. Else returns None.
    let tryRemoveValueAt rowIndex columnIndex sheet : SheetData Option=
        let row = 
            sheet 
            |> getRowAt rowIndex 
            |> Row.tryRemoveCellAt columnIndex
        row
        |> Option.map (fun row ->
            if Row.isEmpty row then
                sheet |> removeRowAt rowIndex
            else
                Row.updateRowSpan row |> ignore
                sheet
        )

    /// Removes the value at the given row- and columnIndex from the sheet.
    let removeValueAt rowIndex columnIndex sheet : SheetData =
        let row = 
            sheet 
            |> getRowAt rowIndex 
            |> Row.removeCellAt columnIndex
        if Row.isEmpty row then
            sheet |> removeRowAt rowIndex
        else
            Row.updateRowSpan row |> ignore
            sheet

    /// Includes a value from sharedStringTable in the cells of the rows of the sheetData
    let includeSharedStringValue (sharedStringTable : SharedStringTable) (sheetData : SheetData) =
        sheetData
        |> mapRows (Row.includeSharedStringValue sharedStringTable)

//----------------------------------------------------------------------------------------------------------------------
//                                                Sheet(s)                                                                 
//----------------------------------------------------------------------------------------------------------------------

    /// Reads the values of all cells from a sheetData and a sharedStringTable and converts them into a sparse matrix. Values are stored sparsely in a dictionary, with the key being a row index and column index tuple.
    let toSparseValueMatrix (sst : SharedStringTable) sheetData =
        let rows = getRows sheetData
        let noOfRows = countRows sheetData
        let noOfCols = 
            // not sure if this is needed since it SEEMS that all left boundaries are always adjusted to the highest value appearing in the rows.
            // if that'd definitely be the case, the left boundary of one of the rows would already be sufficient
            let highestLeftBoundary = 
                rows 
                |> Seq.map (
                    Row.getSpan 
                    >> Row.Spans.toBoundaries
                    >> snd
                )
                |> Seq.max
            int highestLeftBoundary
        let dict = System.Collections.Generic.Dictionary<int * int, string>()
        for iC = 0 to noOfCols - 1 do
            // NOTE: iC is the column index ranging defining the number of columns, columnIndex on the other hand is the real column index as uint
            let columnIndex = iC + 1 |> uint
            for iR = 0 to noOfRows - 1 do
                // NOTE: iR is the row index ranging defining the number of rows, 
                // this rowIndex on the other hand is the real row index since there might be rows not occupied with values (empty rows) in between
                let rowIndex = Seq.item iR rows |> Row.getIndex
                match tryGetCellValueAt (Some sst) rowIndex columnIndex sheetData with
                | Some v -> dict.Add((int rowIndex,int columnIndex),v)
                | None -> ()
        dict