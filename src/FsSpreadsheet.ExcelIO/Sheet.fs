namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging
open System


/// <summary>
/// Part of the Workbook, stores name and other additional info of the sheet. (Unmanaged: Changing a sheet does not alter the associated worksheet which stores the data)
/// </summary>
module Sheet = 

    /// <summary>
    /// Functions for working with Sheets.
    /// </summary>
    // (Unmanaged: changing a sheet does not alter the associated worksheet which stores the data)
    module Sheets = 

        /// <summary>
        /// Creates empty sheets.
        /// </summary>
        let empty () = new Sheets()

        /// <summary>
        /// Returns the first child sheet of the sheets.
        /// </summary>
        let getFirstSheet (sheets : Sheets) = sheets.GetFirstChild<Sheet>()

        /// <summary>
        /// Returns all sheets of the sheets.
        /// </summary>
        let getSheets (sheets : Sheets) = sheets.Elements<Sheet>()

        /// <summary>
        /// Adds a single Sheet to the Sheets.
        /// </summary>
        let addSheet (newSheet : Sheet) (sheets : Sheets) = sheets.AppendChild newSheet
        
        /// <summary>
        /// Adds a Sheet collection to the Sheets.
        /// </summary>
        let addSheets (newSheets : Sheet seq) (sheets : Sheets) = 
            newSheets |> Seq.iter (fun sheet -> addSheet sheet sheets |> ignore) 
            sheets


        //let mapSheets f (sheets:Sheets) = getSheets sheets |> Seq.toArray |> Array.map f
        //let iterSheets f (sheets:Sheets) = getSheets sheets |> Seq.toArray |> Array.iter f
        //let filterSheets f (sheets:Sheets) = 
        //    getSheets sheets |> Seq.toArray |> Array.filter (f >> not)
        //    |> Array.fold (fun st sheet -> removeSheet sheet st) sheets


        /// <summary>
        /// Gets the sheets of the workbook.
        /// </summary>
        let get (workbook : Workbook) = workbook.Sheets

        /// <summary>
        /// Add an empty sheets element to the workbook.
        /// </summary>
        let init (workbook : Workbook) = 
            workbook.AppendChild<Sheets>(Sheets()) |> ignore
            workbook        

        /// <summary>
        /// Returns the existing or a newly created sheets associated with the worksheet.
        /// </summary>
        let getOrInit (workbook : Workbook) =
            if  workbook.Sheets <> null then
                get workbook
            else 
                workbook
                |> init
                |> get        

        let tryGetSheetByName (name : string) (sheets : Sheets) =
            sheets
            |> getSheets
            |> Seq.tryFind (fun sheet -> sheet.Name.Value = name)
    
    /// <summary>
    /// Creates an empty Sheet.
    /// </summary>
    let empty () = Sheet()

    /// <summary>
    /// Sets the name of the sheet (this is the name displayed in MS Excel).
    /// </summary>
    let setName (name : string) (sheet : Sheet) = 
        sheet.Name <- StringValue.FromString name
        sheet

    /// <summary>
    /// Gets the name of the sheet (this is the name displayed in MS Excel).
    /// </summary>
    let getName (sheet : Sheet) = sheet.Name.Value

    /// <summary>
    /// Sets the ID of the sheet (this ID associates the sheet with the worksheet).
    /// </summary>
    let setID id (sheet : Sheet) = 
        sheet.Id <- StringValue.FromString id
        sheet

    /// <summary>
    /// Gets the ID of the sheet (this ID associates the sheet with the worksheet).
    /// </summary>
    let getID (sheet : Sheet) = sheet.Id.Value

    /// <summary>
    /// Sets the SheetID of the sheet (this ID determines the position of the sheet tab in MS Excel).
    /// </summary>
    [<Obsolete("Use setSheetIndex instead ")>]
    let setSheetID id (sheet : Sheet) = 
        sheet.SheetId <- UInt32Value.FromUInt32 id
        sheet

    /// <summary>
    /// Sets the SheetID of the sheet (this ID determines the position of the sheet tab in MS Excel).
    /// </summary>
    let setSheetIndex id (sheet : Sheet) = 
        sheet.SheetId <- UInt32Value.FromUInt32 id
        sheet

    /// <summary>
    /// Gets the SheetID of the sheet (this ID determines the position of the sheet tab in MS Excel).
    /// </summary>
    [<Obsolete("Use getSheetIndex instead.")>]
    let getSheetID (sheet : Sheet) = sheet.SheetId.Value

    /// <summary>
    /// Gets the SheetID of the sheet (this ID determines the position of the sheet tab in MS Excel).
    /// </summary>
    let getSheetIndex (sheet : Sheet) = sheet.SheetId.Value

    /// <summary>
    /// Create a sheet from the id, the name and the sheetID.
    /// </summary>
    let create id name sheetID = 
        Sheet()
        |> setID id
        |> setName name
        |> setSheetID sheetID

    /// <summary>
    /// Returns the item at the given index in the SpreadsheetDocument if it exists. Else returns None.
    /// </summary>
    /// <remarks>SheetIndices are 1-based.</remarks>
    let tryItem (index : uint) (spreadsheetDocument : SpreadsheetDocument) : option<Sheet> = 
        let workbookPart = spreadsheetDocument.WorkbookPart
        workbookPart.Workbook.Descendants<Sheet>()
        |> Seq.tryItem (int index - 1) 

    /// <summary>
    /// Returns the with the given ID in the SpreadsheetDocument if it exists. Else returns None.
    /// </summary>
    let tryGetById (id : string) (spreadsheetDocument : SpreadsheetDocument) : option<Sheet> = 
        let workbookPart = spreadsheetDocument.WorkbookPart
        workbookPart.Workbook.Descendants<Sheet>()
        |> Seq.tryFind (fun s -> s.Id.Value = id) 

    /// <summary>
    /// Returns the item with the given name in the spreadsheetDocument if it exists. Else returns None.
    /// </summary>
    let tryItemByName (name : string) (spreadsheetDocument : SpreadsheetDocument) : option<Sheet> = 
        let workbookPart = spreadsheetDocument.WorkbookPart
        workbookPart.Workbook.Descendants<Sheet>()
        |> Seq.tryFind (fun s -> s.Name.HasValue && s.Name.Value = name)

    /// <summary>
    /// Adds the given sheet to the spreadsheetDocument.
    /// </summary>
    let add (spreadsheetDocument : SpreadsheetDocument) (sheet : Sheet) = 
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        sheets.AppendChild(sheet) |> ignore
        spreadsheetDocument

    /// <summary>
    /// Removes the given sheet from the sheets.
    /// </summary>
    let remove (spreadsheetDocument : SpreadsheetDocument) (sheet : Sheet) =
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        sheets.RemoveChild(sheet) |> ignore
        spreadsheetDocument

    /// <summary>
    /// Returns the sheet for which the predicate returns true (Id Name SheetID -> bool)
    /// </summary>
    let tryFind (predicate:string -> string -> uint32 -> bool) (spreadsheetDocument:SpreadsheetDocument) =
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        Sheets.getSheets sheets
        |> Seq.tryFind (fun sheet -> predicate sheet.Id.Value sheet.Name.Value sheet.SheetId.Value)

    /// <summary>
    /// Counts the number of sheets in the spreadsheetDocument.
    /// </summary>
    let countSheets (spreadsheetDocument : SpreadsheetDocument) =
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        Sheets.getSheets sheets |> Seq.length


