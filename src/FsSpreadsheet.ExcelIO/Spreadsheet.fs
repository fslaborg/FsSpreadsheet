namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Spreadsheet

open FsSpreadsheet

/// <summary>
/// Functions for working the spreadsheet document.
/// </summary>
module Spreadsheet = 

    /// <summary>
    /// Opens the spreadsheet located at the given path and initialized a FileStream.
    /// </summary>
    let fromFile (path : string) isEditable = SpreadsheetDocument.Open(path,isEditable)

    /// <summary>
    /// Opens the spreadsheet from the given FileStream.
    /// </summary>
    let fromStream (stream : System.IO.Stream) isEditable = SpreadsheetDocument.Open(stream,isEditable)

    /// <summary>
    /// Initializes a new empty spreadsheet at the given path.
    /// </summary>
    let initEmpty (path : string) = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook)

    /// <summary>
    /// Initializes a new empty spreadsheet in the given stream.
    /// </summary>
    let initEmptyOnStream (stream : System.IO.Stream) = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)

    /// <summary>
    /// Gets the workbookPart of the spreadsheet.
    /// </summary>
    let getWorkbookPart (spreadsheet : SpreadsheetDocument) = spreadsheet.WorkbookPart

    /// <summary>
    /// Initialized a new workbookPart in the spreadsheetDocument but only if there is none.
    /// </summary>
    // Only if none there
    let initWorkbookPart (spreadsheet : SpreadsheetDocument) = spreadsheet.AddWorkbookPart()

    /// <summary>
    /// Saves changes made to the spreadsheet.
    /// </summary>
    let saveChanges (spreadsheet : SpreadsheetDocument) = 
        spreadsheet.Save() 
        spreadsheet

    /// <summary>
    /// Closes the FileStream to the spreadsheet.
    /// </summary>
    let close (spreadsheet : SpreadsheetDocument) = spreadsheet.Close()

    /// <summary>
    /// Saves changes made to the spreadsheet to the given path.
    /// </summary>
    let saveAs path (spreadsheet : SpreadsheetDocument) = 
        spreadsheet.SaveAs(path) :?> SpreadsheetDocument
        |> close
        spreadsheet

    /// <summary>
    /// Initializes a new spreadsheet with an empty sheet at the given path.
    /// </summary>
    let init sheetName (path : string) = 
        let doc = initEmpty path
        let workbookPart = initWorkbookPart doc

        WorkbookPart.appendSheet sheetName (SheetData.empty ()) workbookPart |> ignore
        doc

    /// <summary>
    /// Initializes a new spreadsheet with an empty sheet in the given stream.
    /// </summary>
    let initOnStream sheetName (stream : System.IO.Stream) = 
        let doc = initEmptyOnStream stream
        let workbookPart = initWorkbookPart doc

        WorkbookPart.appendSheet sheetName (SheetData.empty ()) workbookPart |> ignore
        doc

    /// <summary>
    /// Initializes a new spreadsheet with an empty sheet and a sharedStringTable at the given path.
    /// </summary>
    let initWithSst sheetName (path : string) = 
        let doc = init sheetName path
        let workbookPart = getWorkbookPart doc

        let sharedStringTablePart = WorkbookPart.getOrInitSharedStringTablePart workbookPart
        SharedStringTable.init sharedStringTablePart |> ignore

        doc

    /// <summary>
    /// Initializes a new spreadsheet with an empty sheet and a sharedStringTable in the given stream.
    /// </summary>
    let initWithSstOnStream sheetName (stream : System.IO.Stream) = 
        let doc = initOnStream sheetName stream
        let workbookPart = getWorkbookPart doc

        let sharedStringTablePart = WorkbookPart.getOrInitSharedStringTablePart workbookPart
        SharedStringTable.init sharedStringTablePart |> ignore

        doc

    /// <summary>
    /// Gets the sharedStringTable of a spreadsheet.
    /// </summary>
    let getSharedStringTable (spreadsheetDocument : SpreadsheetDocument) =
        spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable

    /// <summary>
    /// Gets the sharedStringTable of the spreadsheet if it exists, else returns None.
    /// </summary>
    let tryGetSharedStringTable (spreadsheetDocument : SpreadsheetDocument) =
        try spreadsheetDocument.WorkbookPart.SharedStringTablePart.SharedStringTable |> Some
        with | _ -> None

    /// <summary>
    /// Gets the sharedStringTablePart. If it does not exist, creates a new one.
    /// </summary>
    let getOrInitSharedStringTablePart (spreadsheetDocument : SpreadsheetDocument) =
        let workbookPart = spreadsheetDocument.WorkbookPart    
        let sstp = workbookPart.GetPartsOfType<SharedStringTablePart>()
        match sstp |> Seq.tryHead with
        | Some sst -> sst
        | None -> workbookPart.AddNewPart<SharedStringTablePart>()

    /// <summary>
    /// Returns the worksheetPart associated to the sheet with the given name.
    /// </summary>
    let tryGetWorksheetPartBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        Sheet.tryItemByName name spreadsheetDocument
        |> Option.map (fun sheet -> 
            spreadsheetDocument.WorkbookPart
            |> Worksheet.WorksheetPart.getByID sheet.Id.Value 
        )      

    /// <summary>
    /// Returns the worksheetPart for the given 0-based sheetIndex of the given spreadsheetDocument. 
    /// </summary>
    let tryGetWorksheetPartBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        Sheet.tryItem sheetIndex spreadsheetDocument
        |> Option.map (fun sheet -> 
            spreadsheetDocument.WorkbookPart
            |> Worksheet.WorksheetPart.getByID sheet.Id.Value 
        )   
        
    /// <summary>
    /// Returns the sheetData for the given 0-based sheetIndex of the given spreadsheetDocument if it exists. Else returns None.
    /// </summary>
    let tryGetSheetBySheetName (name : string) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetName name spreadsheetDocument
        |> Option.map (Worksheet.get >> Worksheet.getSheetData)

    /// <summary>
    /// Returns the sheetData for the given 0-based sheetIndex of the given spreadsheetDocument. 
    /// </summary>
    let tryGetSheetBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        tryGetWorksheetPartBySheetIndex sheetIndex spreadsheetDocument
        |> Option.map (Worksheet.get >> Worksheet.getSheetData)
        

    /// <summary>
    /// Returns a sequence of rows containing the cells for the given 0-based sheetIndex of the given spreadsheetDocument. 
    /// </summary>
    let getRowsBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =

        match (Sheet.tryItem sheetIndex spreadsheetDocument) with
        | Some (sheet) ->
            let workbookPart = spreadsheetDocument.WorkbookPart
            let worksheetPart = Worksheet.WorksheetPart.getByID sheet.Id.Value workbookPart     
            let stringTablePart = getOrInitSharedStringTablePart spreadsheetDocument
            let sst = SharedStringTable.toSST stringTablePart.SharedStringTable
            seq {
            use reader = OpenXmlReader.Create(worksheetPart)
      
            while reader.Read() do
                if (reader.ElementType = typeof<Row>) then 
                    let row = reader.LoadCurrentElement() :?> Row
                    row.Elements()
                    |> Seq.iter (fun item -> 
                        let cell = item :?> Cell
                        Cell.includeSharedStringValue sst cell |> ignore
                        )
                    yield row 
            }
        | None -> Seq.empty

    /// <summary>
    /// Returns a 1D-sequence of Cells for the given Sheet of the given SpreadsheetDocument.
    /// </summary>
    let getCellsBySheet (sheet : Sheet) (spreadsheetDocument : SpreadsheetDocument) =
        let workbookPart = spreadsheetDocument.WorkbookPart
        let worksheetPart = Worksheet.WorksheetPart.getByID sheet.Id.Value workbookPart
        
        let includeSSV = 
            match tryGetSharedStringTable spreadsheetDocument with 
            | Some sst -> 
                let sstArray = sst |> SharedStringTable.toSST
                Cell.includeSharedStringValue sstArray
            | None -> id
        seq {
        use reader = OpenXmlReader.Create(worksheetPart)
        
        while reader.Read() do
            if (reader.ElementType = typeof<Cell>) then 
                let cell    = reader.LoadCurrentElement() :?> Cell 
                let cellRef = if cell.CellReference.HasValue then cell.CellReference.Value else ""
                yield includeSSV cell
        }

    /// <summary>
    /// Returns a 1D-sequence of Cells for the given sheetIndex of the given SpreadsheetDocument.
    /// </summary>
    /// <remarks>SheetIndices are 1-based.</remarks>
    let getCellsBySheetIndex (sheetIndex : uint) (spreadsheetDocument : SpreadsheetDocument) =
        match Sheet.tryItem sheetIndex spreadsheetDocument with
        | Some sheet -> getCellsBySheet sheet spreadsheetDocument
        | None -> seq {()}

    /// <summary>
    /// Returns a 1D-sequence of Cells for the given sheetIndex of the given SpreadsheetDocument.
    /// </summary>
    /// <remarks>SheetIndices are 1-based.</remarks>
    let getCellsBySheetID (sheetID : string) (spreadsheetDocument : SpreadsheetDocument) =
        match Sheet.tryGetById sheetID spreadsheetDocument with
        | Some sheet -> getCellsBySheet sheet spreadsheetDocument
        | None -> seq {()}

    //----------------------------------------------------------------------------------------------------------------------
    //                                      High level functions                                                            
    //----------------------------------------------------------------------------------------------------------------------

    //Rows

    /// <summary>
    /// </summary>
    let mapRowOfSheet (sheetId) (rowId) (rowF: Row -> Row) : SpreadsheetDocument = 
        //get workbook part
        //get sheet data by sheetId
        //get row at rowId
        //apply rowF to row and update 
        //return updated doc
        raise (System.NotImplementedException())

    /// <summary>
    /// </summary>
    let mapRowsOfSheet (sheetId) (rowF: Row -> Row) : SpreadsheetDocument = raise (System.NotImplementedException())

    /// <summary>
    /// </summary>
    let appendRowValuesToSheet (sheetId) (rowValues: seq<'T>) : SpreadsheetDocument = raise (System.NotImplementedException())

    /// <summary>
    /// </summary>
    let insertRowValuesIntoSheetAt (sheetId) (rowId) (rowValues: seq<'T>) : SpreadsheetDocument = raise (System.NotImplementedException())

    /// <summary>
    /// </summary>
    let insertValueIntoSheetAt (sheetId) (rowId) (colId) (value: 'T) : SpreadsheetDocument = raise (System.NotImplementedException())

    /// <summary>
    /// </summary>
    let setValueInSheetAt (sheetId) (rowId) (colId) (value: 'T) : SpreadsheetDocument = raise (System.NotImplementedException())

    /// <summary>
    /// </summary>
    let deleteRowFromSheet (sheetId) (rowId) : SpreadsheetDocument = raise (System.NotImplementedException())

    //...