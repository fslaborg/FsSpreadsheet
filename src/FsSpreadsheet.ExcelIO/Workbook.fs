namespace FsSpreadsheet.ExcelIO


open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging

/// Functions for manipulating Workbooks. (Unmanaged: changing the sheets does not alter the associated worksheets which store the data)
module Workbook =

    /// Creates an empty Workbook.
    let empty () = new Workbook()
    
    /// Gets the Workbook of the WorkbookPart.
    let get (workbookPart : WorkbookPart) = workbookPart.Workbook 

    /// Sets the Workbook of the WorkbookPart.
    let set (workbook : Workbook) (workbookPart : WorkbookPart) = 
        workbookPart.Workbook <- workbook
        workbookPart

    /// Sets an empty Workbook in the given WorkbookPart.
    let init (workbookPart : WorkbookPart) = 
        set (Workbook()) workbookPart

    /// Returns the Workbookpart associated with the existing or a newly created Workbook.
    let getOrInit (workbookPart : WorkbookPart) =
        if workbookPart.Workbook <> null then
            get workbookPart
        else 
            workbookPart
            |> init
            |> get

    //// Adds sheet to workbook
    //let addSheet (sheet : Sheet) (workbook:Workbook) =
    //    let sheets = Sheet.Sheets.getOrInit workbook
    //    Sheet.Sheets.addSheet sheet sheets |> ignore
    //    workbook

/// Functions for working with WorkbookParts.
module WorkbookPart = 

    /// Add a WorksheetPart to the WorkbookPart.
    let addWorksheetPart (worksheetPart : WorksheetPart) (workbookPart : WorkbookPart) = 
        workbookPart.AddPart(worksheetPart)

    /// Add an empty WorksheetPart to the WorkbookPart.
    let initWorksheetPart (workbookPart : WorkbookPart) = workbookPart.AddNewPart<WorksheetPart>()

    /// Get the WorksheetParts of the WorkbookPart.
    let getWorkSheetParts (workbookPart : WorkbookPart) = workbookPart.WorksheetParts

    /// Returns true if the WorkbookPart contains at least one WorksheetPart.
    let containsWorkSheetParts (workbookPart : WorkbookPart) = workbookPart.GetPartsOfType<WorksheetPart>() |> Seq.length |> (<>) 0

    /// Gets the WorksheetPart of the WorkbookPart with the given ID.
    let getWorksheetPartById (id : string) (workbookPart : WorkbookPart) = workbookPart.GetPartById(id) :?> WorksheetPart 

    /// If the WorkbookPart contains the WorksheetPart with the given ID, returns it. Else returns None.
    let tryGetWorksheetPartById (id : string) (workbookPart : WorkbookPart) = 
        try workbookPart.GetPartById(id) :?> WorksheetPart  |> Some with
        | _ -> None

    /// Gets the ID of the WorksheetPart of the WorkbookPart.
    let getWorksheetPartID (worksheetPart : WorksheetPart) (workbookPart : WorkbookPart) = workbookPart.GetIdOfPart worksheetPart
    //let addworkSheet (workbookPart:WorkbookPart) (worksheet : Worksheet) = 
    //    let emptySheet = (addNewWorksheetPart workbookPart)
    //    emptySheet.Worksheet <- worksheet

    /// Gets the SharedStringTablePart of a WorkbookPart.
    let getSharedStringTablePart (workbookPart : WorkbookPart) = workbookPart.SharedStringTablePart
    
    /// Sets an empty SharedStringTablePart in the given WorkbookPart.
    let initSharedStringTablePart (workbookPart : WorkbookPart) = 
        workbookPart.AddNewPart<SharedStringTablePart>() |> ignore
        workbookPart

    /// Returns true if the WorkbookPart contains a SharedStringTablePart.
    let containsSharedStringTablePart (workbookPart : WorkbookPart) = workbookPart.SharedStringTablePart <> null

    /// Returns the existing or a newly created SharedStringTablePart associated with the WorkbookPart.
    let getOrInitSharedStringTablePart (workbookPart : WorkbookPart) =
        if containsSharedStringTablePart workbookPart then
            getSharedStringTablePart workbookPart
        else 
            initSharedStringTablePart workbookPart
            |> getSharedStringTablePart
    
    /// Returns the SharedStringTable of a WorkbookPart.
    let getSharedStringTable (workbookPart : WorkbookPart) =
        workbookPart 
        |> getSharedStringTablePart 
        |> SharedStringTable.get

    /// Returns the SheetData of the first sheet of the given WorkbookPart.
    let getDataOfFirstSheet (workbookPart : WorkbookPart) = 
        workbookPart
        |> getWorkSheetParts
        |> Seq.head
        |> Worksheet.get
        |> Worksheet.getSheetData

    /// Appends a new sheet with the given SheetData to the SpreadsheetDocument.
    // to-do: guard if sheet of name already exists
    let appendSheet (sheetName : string) (data : SheetData) (workbookPart : WorkbookPart) =

        let workbook = Workbook.getOrInit  workbookPart

        let worksheetPart = initWorksheetPart workbookPart

        Worksheet.getOrInit worksheetPart
        |> Worksheet.addSheetData data
        |> ignore
        
        let sheets = Sheet.Sheets.getOrInit workbook
        let id = getWorksheetPartID worksheetPart workbookPart
        let sheetID = 
            sheets |> Sheet.Sheets.getSheets |> Seq.map Sheet.getSheetID
            |> fun s -> 
                if Seq.length s = 0 then 1u
                else s |> Seq.max |> (+) 1ul

        let sheet = Sheet.create id sheetName sheetID

        sheets.AppendChild(sheet) |> ignore
        workbookPart

    /// Replaces the SheetData of the Sheet with the given sheetname.
    let replaceSheetDataByName (sheetName : string) (data : SheetData) (workbookPart : WorkbookPart) =

        let workbook = Workbook.getOrInit  workbookPart
    
        let sheets = Sheet.Sheets.getOrInit workbook
        let id = 
            sheets |> Sheet.Sheets.getSheets
            |> Seq.find (fun sheet -> Sheet.getName sheet = sheetName)
            |> Sheet.getID

        getWorksheetPartById id workbookPart
        |> Worksheet.getOrInit
        |> Worksheet.setSheetData data
        |> ignore 

        workbookPart
