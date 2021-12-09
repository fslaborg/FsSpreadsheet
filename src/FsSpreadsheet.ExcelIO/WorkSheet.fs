namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging

open FsSpreadsheet

/// Stores data of the sheet and the index of the sheet and
/// functions for working with the worksheetpart. (Unmanaged: changing a worksheet does not alter the sheet which links the worksheet to the excel workbook)
module Worksheet = 

    /// Empty Worksheet
    let empty() = Worksheet()

    /// Associates a SheetData with the Worksheet.
    let addSheetData (sheetData : SheetData) (worksheet : Worksheet) = 
        let children = worksheet.ChildElements
        let posPageMargins =
            children
            |> Seq.tryFindIndex (fun c -> c.LocalName = "pageMargins")
        match posPageMargins with
        | Some pos  -> worksheet.InsertAt(sheetData, pos)
        | None      -> worksheet.AppendChild(sheetData)
        |> ignore
        worksheet

    /// Returns true, if the Worksheet contains SheetData.
    let hasSheetData (worksheet : Worksheet) = 
        worksheet.HasChildren

    /// Creates a Worksheet containing the given SheetData.
    let ofSheetData (sheetData : SheetData) = 
        Worksheet(sheetData)

    /// Returns the sheetdata associated with the Worksheet.
    let getSheetData (worksheet : Worksheet) = 
        worksheet.GetFirstChild<SheetData>()
      
    /// Sets the SheetData of a Worksheet.
    let setSheetData (sheetData : SheetData) (worksheet : Worksheet) =
        if hasSheetData worksheet then
            worksheet.RemoveChild(getSheetData worksheet)
            |> ignore
        addSheetData sheetData worksheet

    // Returns the Worksheet associated with the WorksheetPart.
    let get (worksheetPart : WorksheetPart) = 
        worksheetPart.Worksheet

    /// Sets the given Worksheet with the WorksheetPart.
    let setWorksheet (worksheet : Worksheet) (worksheetPart : WorksheetPart) = 
        worksheetPart.Worksheet <- worksheet
        worksheetPart

    /// Associates an empty Worksheet with the WworksheetPart.
    let init (worksheetPart : WorksheetPart) = 
        worksheetPart
        |> setWorksheet (empty())

    /// Returns the existing or a newly created Worksheet associated with the WorksheetPart.
    let getOrInit (worksheetPart : WorksheetPart) =
        if worksheetPart.Worksheet <> null then
            get worksheetPart
        else 
            worksheetPart
            |> init
            |> get

    /// Functions for extracting / working with WorksheetParts.
    module WorksheetPart = 

        /// Returns the WorksheetPart matching the given sheetID.
        let getByID sheetID (workbookPart : WorkbookPart) = 
            workbookPart.GetPartById(sheetID) :?> WorksheetPart  
            
        /// Returns the SheetData associated with the WorksheetPart.
        let getSheetData (worksheetPart : WorksheetPart) =
            get worksheetPart |> getSheetData

        /// Returns the WorksheetCommentsPart associated with a WorksheetPart.
        let getWorksheetCommentsPart (worksheetPart : WorksheetPart) = worksheetPart.WorksheetCommentsPart

    /// Functions for extracting / working with WorksheetCommentsParts.
    module WorksheetCommentsPart =
        
        /// Returns the WorksheetCommentsPart associated with a WorksheetPart.
        let get (worksheetPart : WorksheetPart) = WorksheetPart.getWorksheetCommentsPart worksheetPart
        
        /// Returns the comments of the WorksheetCommentsPart.
        let getComments (worksheetCommentsPart : WorksheetCommentsPart) = worksheetCommentsPart.Comments

    // TO DO: Atm. both types of comments (REAL comments and notes) are mixed. They seem to only differ in terms of their text: Comments have a disclaimer like "Comment:" or "Reply:" (the latter if it's a reply to a comment) while notes do not have that BUT have text formatting (can be seen in comments.xml in .xlsx archives))
    /// Functions for working with Comments.
    module Comments =
        
        /// Returns the comments of the WorksheetCommentsPart.
        let get (worksheetCommentsPart : WorksheetCommentsPart) = worksheetCommentsPart.Comments

        /// Returns the CommentList of the given Comments.
        let getCommentList (comments : Comments) = comments.CommentList

        /// <summary>Returns a sequence of author names from the Comments.</summary>
        /// <remarks>Author names might be encrypted in the pattern of <code>tc={...}</code></remarks>
        let getAuthors (comments : Comments) = comments.Authors |> Seq.map (fun a -> a.InnerText)

        /// Returns all Comments and Notes as strings of a CommentList.
        let getCommentAndNoteTexts (commentList : CommentList) = 
            commentList |> Seq.map (fun c -> c.InnerText)

        /// Returns a triple of Comments consisting of the author, the comment text written, and the cell reference (A1-style).
        let getCommentsAuthorsTextsRefs (comments : Comments) =
            let authors = comments.Authors.Elements<DocumentFormat.OpenXml.Spreadsheet.Author>()
            let refsAuthorsTexts = 
                comments.CommentList.Elements<Comment>() 
                |> Seq.choose (
                    fun c ->
                        match c.CommentText.Text with
                        | null -> None
                        | _ -> 
                            Some (
                                c.Reference.Value,
                                (Seq.item (int c.AuthorId.Value) authors).Text,
                                c.CommentText.Text.InnerText
                            )
                )
            refsAuthorsTexts

        /// Returns a triple of Notes consisting of the author, the note text written, and the cell reference (A1-style).
        let getNotesAuthorsTextsRefs (comments : Comments) =
            let authors = comments.Authors.Elements<DocumentFormat.OpenXml.Spreadsheet.Author>()
            let refsAuthorsTexts = 
                comments.CommentList.Elements<Comment>() 
                |> Seq.choose (
                    fun c ->
                        match c.CommentText.Text with
                        | null -> 
                            Some (
                                c.Reference.Value,
                                (Seq.item (int c.AuthorId.Value) authors).Text,
                                c.CommentText.Text.InnerText
                            )

                        | _ -> None
                )
            refsAuthorsTexts


    //let insertCellData (cell:CellData.CellDataValue) (worksheet : Worksheet) =
        
    ///Convenience

    //let insertRow (rowIndex) (values: 'T seq) (worksheet:Worksheet) = notImplemented()
    //let overWriteRow (rowIndex) (values: 'T seq) (worksheet:Worksheet) = notImplemented()
    //let appendRow (values: 'T seq) (worksheet:Worksheet) = notImplemented()
    //let getRow (rowIndex) (worksheet:Worksheet) = notImplemented()
    //let deleteRow rowIndex (worksheet:Worksheet) = notImplemented()

    //let insertColumn (columnIndex) (values: 'T seq) (worksheet:Worksheet) = notImplemented()
    //let overWriteColumn (columnIndex) (values: 'T seq) (worksheet:Worksheet) = notImplemented()
    //let appendColumn (values: 'T seq) (worksheet:Worksheet) = notImplemented()
    //let getColumn (columnIndex) (worksheet:Worksheet) = notImplemented()
    //let deleteColumn (columnIndex) (worksheet:Worksheet) = notImplemented()

    ////let setCellValue (rowIndex,columnIndex) value (worksheet:Worksheet) = notImplemented()
    //let setCellValue adress value (worksheet:Worksheet) = notImplemented()
    //let inferCellValue adress (worksheet:Worksheet) = notImplemented()
    //let deleteCellValue adress (worksheet:Worksheet) = notImplemented()



    //let setID id (worksheetPart : WorksheetPart) = notImplemented()
    //let getID (worksheetPart : WorksheetPart) = notImplemented()
