namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

open FsSpreadsheet

/// Functions for working the spreadsheet document.
module Spreadsheet = 

    /// Opens the spreadsheet located at the given path and initialized a FileStream.
    let fromFile (path : string) isEditable = SpreadsheetDocument.Open(path,isEditable)

    /// Opens the spreadsheet from the given FileStream.
    let fromStream (stream : System.IO.Stream) isEditable = SpreadsheetDocument.Open(stream,isEditable)

    /// Initializes a new empty spreadsheet at the given path.
    let initEmpty (path : string) = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)

    /// Initializes a new empty spreadsheet in the given stream.
    let initEmptyOnStream (stream : System.IO.Stream) = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)

    // Gets the workbookPart of the spreadsheet.
    let getWorkbookPart (spreadsheet : SpreadsheetDocument) = spreadsheet.WorkbookPart

    // Only if none there
    /// Initialized a new workbookPart in the spreadsheetDocument but only if there is none.
    let initWorkbookPart (spreadsheet : SpreadsheetDocument) = spreadsheet.AddWorkbookPart()

    /// Saves changes made to the spreadsheet.
    let saveChanges (spreadsheet : SpreadsheetDocument) = 
        spreadsheet.Save() 
        spreadsheet

    /// Closes the FileStream to the spreadsheet.
    let close (spreadsheet : SpreadsheetDocument) = spreadsheet.Close()

    /// Saves changes made to the spreadsheet to the given path.
    let saveAs path (spreadsheet : SpreadsheetDocument) = 
        spreadsheet.SaveAs(path) :?> SpreadsheetDocument
        |> close
        spreadsheet

    /// Initializes a new spreadsheet with an empty sheet at the given path.
    let init sheetName (path : string) = 
        let doc = initEmpty path
        let workbookPart = initWorkbookPart doc

        WorkbookPart.appendSheet sheetName (SheetData.empty ()) workbookPart |> ignore
        doc

    /// Initializes a new spreadsheet with an empty sheet in the given stream.
    let initOnStream sheetName (stream : System.IO.Stream) = 
        let doc = initEmptyOnStream stream
        let workbookPart = initWorkbookPart doc

        WorkbookPart.appendSheet sheetName (SheetData.empty ()) workbookPart |> ignore
        doc

    /// Initializes a new spreadsheet with an empty sheet and a sharedStringTable at the given path.
    let initWithSst sheetName (path : string) = 
        let doc = init sheetName path
        let workbookPart = getWorkbookPart doc

        let sharedStringTablePart = WorkbookPart.getOrInitSharedStringTablePart workbookPart
        SharedStringTable.init sharedStringTablePart |> ignore

        doc

    /// Initializes a new spreadsheet with an empty sheet and a sharedStringTable in the given stream.
    let initWithSstOnStream sheetName (stream : System.IO.Stream) = 
        let doc = initOnStream sheetName stream
        let workbookPart = getWorkbookPart doc

        let sharedStringTablePart = WorkbookPart.getOrInitSharedStringTablePart workbookPart
        SharedStringTable.init sharedStringTablePart |> ignore

        doc

    // Gets the sharedStringTable of a spreadsheet.
    let getSharedStringTable (spreadsheetDocument : SpreadsheetDocument) =
        spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable

    // Gets the sharedStringTable of the spreadsheet if it exists, else returns None.
    let tryGetSharedStringTable (spreadsheetDocument : SpreadsheetDocument) =
        try spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable |> Some
        with | _ -> None

    // Gets the sharedStringTablePart. If it does not exist, creates a new one.
    let getOrInitSharedStringTablePart (spreadsheetDocument : SpreadsheetDocument) =
        let workbookPart = spreadsheetDocument.WorkbookPart    
        let sstp = workbookPart.GetPartsOfType<SharedStringTablePart>()
        match sstp |> Seq.tryHead with
        | Some sst -> sst
        | None -> workbookPart.AddNewPart<SharedStringTablePart>()

    /// Returns the worksheetPart associated to the sheet with the given name.
    let tryGetWorksheetPartBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        Sheet.tryItemByName name spreadsheetDocument
        |> Option.map (fun sheet -> 
            spreadsheetDocument.WorkbookPart
            |> Worksheet.WorksheetPart.getByID sheet.Id.Value 
        )      

    /// Returns the worksheetPart for the given 0-based sheetIndex of the given spreadsheetDocument. 
    let tryGetWorksheetPartBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        Sheet.tryItem sheetIndex spreadsheetDocument
        |> Option.map (fun sheet -> 
            spreadsheetDocument.WorkbookPart
            |> Worksheet.WorksheetPart.getByID sheet.Id.Value 
        )   
        
    /// Returns the sheetData for the given 0-based sheetIndex of the given spreadsheetDocument if it exists. Else returns None.
    let tryGetSheetBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetName name spreadsheetDocument
        |> Option.map (Worksheet.get >> Worksheet.getSheetData)

    /// Returns the sheetData for the given 0-based sheetIndex of the given spreadsheetDocument. 
    let tryGetSheetBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetIndex sheetIndex spreadsheetDocument
        |> Option.map (Worksheet.get >> Worksheet.getSheetData)
        

    /// Returns a sequence of rows containing the cells for the given 0-based sheetIndex of the given spreadsheetDocument. 
    let getRowsBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =

        match (Sheet.tryItem sheetIndex spreadsheetDocument) with
        | Some (sheet) ->
            let workbookPart = spreadsheetDocument.WorkbookPart
            let worksheetPart = Worksheet.WorksheetPart.getByID sheet.Id.Value workbookPart     
            let stringTablePart = getOrInitSharedStringTablePart spreadsheetDocument
            seq {
            use reader = OpenXmlReader.Create(worksheetPart)
      
            while reader.Read() do
                if (reader.ElementType = typeof<Row>) then 
                    let row = reader.LoadCurrentElement() :?> Row
                    row.Elements()
                    |> Seq.iter (fun item -> 
                        let cell = item :?> Cell
                        Cell.includeSharedStringValue stringTablePart.SharedStringTable cell |> ignore
                        )
                    yield row 
            }
        | None -> Seq.empty

    /// <summary>Returns a 1D-sequence of Cells for the given sheetIndex of the given SpreadsheetDocument.</summary>
    /// <remarks>SheetIndices are 1-based.</remarks>
    let getCellsBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =

        match (Sheet.tryItem sheetIndex spreadsheetDocument) with
        | Some (sheet) ->
            let workbookPart = spreadsheetDocument.WorkbookPart
            let worksheetPart = Worksheet.WorksheetPart.getByID sheet.Id.Value workbookPart
            let stringTablePart = getOrInitSharedStringTablePart spreadsheetDocument
            seq {
            use reader = OpenXmlReader.Create(worksheetPart)
        
            while reader.Read() do
                if (reader.ElementType = typeof<Cell>) then 
                    let cell    = reader.LoadCurrentElement() :?> Cell 
                    let cellRef = if cell.CellReference.HasValue then cell.CellReference.Value else ""
                    yield Cell.includeSharedStringValue stringTablePart.SharedStringTable cell
            }
        | None -> seq {()}

    //----------------------------------------------------------------------------------------------------------------------
    //                                      High level functions                                                            
    //----------------------------------------------------------------------------------------------------------------------

    //Rows

    let mapRowOfSheet (sheetId) (rowId) (rowF: Row -> Row) : SpreadsheetDocument = 
        //get workbook part
        //get sheet data by sheetId
        //get row at rowId
        //apply rowF to row and update 
        //return updated doc
        raise (System.NotImplementedException())

    let mapRowsOfSheet (sheetId) (rowF: Row -> Row) : SpreadsheetDocument = raise (System.NotImplementedException())

    let appendRowValuesToSheet (sheetId) (rowValues: seq<'T>) : SpreadsheetDocument = raise (System.NotImplementedException())

    let insertRowValuesIntoSheetAt (sheetId) (rowId) (rowValues: seq<'T>) : SpreadsheetDocument = raise (System.NotImplementedException())

    let insertValueIntoSheetAt (sheetId) (rowId) (colId) (value: 'T) : SpreadsheetDocument = raise (System.NotImplementedException())

    let setValueInSheetAt (sheetId) (rowId) (colId) (value: 'T) : SpreadsheetDocument = raise (System.NotImplementedException())

    let deleteRowFromSheet (sheetId) (rowId) : SpreadsheetDocument = raise (System.NotImplementedException())

    //...