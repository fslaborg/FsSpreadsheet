namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging


/// Part of the Workbook, stores name and other additional info of the sheet. (Unmanaged: Changing a sheet does not alter the associated worksheet which stores the data)
module Sheet = 
    
    /// Functions for working with Sheets. (Unmanaged: changing a sheet does not alter the associated worksheet which stores the data)
    module Sheets = 

        /// Creates empty sheets.
        let empty () = new Sheets()

        /// Returns the first child sheet of the sheets.
        let getFirstSheet (sheets : Sheets) = sheets.GetFirstChild<Sheet>()

        /// Returns all sheets of the sheets.
        let getSheets (sheets : Sheets) = sheets.Elements<Sheet>()

        /// Adds a single Sheet to the Sheets.
        let addSheet (newSheet : Sheet) (sheets : Sheets) = sheets.AppendChild newSheet
        
        /// Adds a Sheet collection to the Sheets.
        let addSheets (newSheets : Sheet seq) (sheets : Sheets) = 
            newSheets |> Seq.iter (fun sheet -> addSheet sheet sheets |> ignore) 
            sheets


        //let mapSheets f (sheets:Sheets) = getSheets sheets |> Seq.toArray |> Array.map f
        //let iterSheets f (sheets:Sheets) = getSheets sheets |> Seq.toArray |> Array.iter f
        //let filterSheets f (sheets:Sheets) = 
        //    getSheets sheets |> Seq.toArray |> Array.filter (f >> not)
        //    |> Array.fold (fun st sheet -> removeSheet sheet st) sheets


        /// Gets the sheets of the workbook.
        let get (workbook : Workbook) = workbook.Sheets

        /// Add an empty sheets element to the workbook.
        let init (workbook : Workbook) = 
            workbook.AppendChild<Sheets>(Sheets()) |> ignore
            workbook        

        /// Returns the existing or a newly created sheets associated with the worksheet.
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
    
    /// Creates an empty Sheet.
    let empty () = Sheet()

    /// Sets the name of the sheet (this is the name displayed in MS Excel).
    let setName (name : string) (sheet : Sheet) = 
        sheet.Name <- StringValue.FromString name
        sheet

    /// Gets the name of the sheet (this is the name displayed in MS Excel).
    let getName (sheet : Sheet) = sheet.Name.Value

    /// Sets the ID of the sheet (this ID associates the sheet with the worksheet).
    let setID id (sheet : Sheet) = 
        sheet.Id <- StringValue.FromString id
        sheet

    /// Gets the ID of the sheet (this ID associates the sheet with the worksheet).
    let getID (sheet : Sheet) = sheet.Id.Value

    /// Sets the SheetID of the sheet (this ID determines the position of the sheet tab in MS Excel).
    let setSheetID id (sheet : Sheet) = 
        sheet.SheetId <- UInt32Value.FromUInt32 id
        sheet

    /// Gets the SheetID of the sheet (this ID determines the position of the sheet tab in MS Excel).
    let getSheetID (sheet : Sheet) = sheet.SheetId.Value

    /// Create a sheet from the id, the name and the sheetID.
    let create id name sheetID = 
        Sheet()
        |> setID id
        |> setName name
        |> setSheetID sheetID

    /// Returns the item of the given index in the spreadsheetDocument if it exists. Else returns None.
    let tryItem (index : uint) (spreadsheetDocument : SpreadsheetDocument) : option<Sheet> = 
        let workbookPart = spreadsheetDocument.WorkbookPart    
        workbookPart.Workbook.Descendants<Sheet>()
        |> Seq.tryItem (int index) 

    /// Returns the item with the given name in the spreadsheetDocument if it exists. Else returns None.
    let tryItemByName (name : string) (spreadsheetDocument : SpreadsheetDocument) : option<Sheet> = 
        let workbookPart = spreadsheetDocument.WorkbookPart    
        workbookPart.Workbook.Descendants<Sheet>()
        |> Seq.tryFind (fun s -> s.Name.HasValue && s.Name.Value = name)

    /// Adds the given sheet to the spreadsheetDocument.
    let add (spreadsheetDocument : SpreadsheetDocument) (sheet : Sheet) = 
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        sheets.AppendChild(sheet) |> ignore
        spreadsheetDocument

    /// Removes the given sheet from the sheets.
    let remove (spreadsheetDocument : SpreadsheetDocument) (sheet : Sheet) =
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        sheets.RemoveChild(sheet) |> ignore
        spreadsheetDocument

    /// Returns the sheet for which the predicate returns true (Id Name SheetID -> bool)
    let tryFind (predicate:string -> string -> uint32 -> bool) (spreadsheetDocument:SpreadsheetDocument) =
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        Sheets.getSheets sheets
        |> Seq.tryFind (fun sheet -> predicate sheet.Id.Value sheet.Name.Value sheet.SheetId.Value)

    /// Counts the number of sheets in the spreadsheetDocument.
    let countSheets (spreadsheetDocument : SpreadsheetDocument) =
        let sheets = spreadsheetDocument.WorkbookPart.Workbook.Sheets
        Sheets.getSheets sheets |> Seq.length


