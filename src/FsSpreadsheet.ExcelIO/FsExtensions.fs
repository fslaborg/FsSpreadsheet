namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open FsSpreadsheet
open System.IO

/// <summary>
/// Classes that extend the core FsSpreadsheet library with IO functionalities.
/// </summary>
[<AutoOpen>]
module FsExtensions =


    type DataType with

        /// <summary>
        /// Converts a given CellValues to the respective DataType.
        /// </summary>
        static member ofXlsXCell (doc : Packaging.SpreadsheetDocument)  (cell : Cell) =

            //https://stackoverflow.com/a/13178043/12858021
            //https://stackoverflow.com/a/55425719/12858021
            // if styleindex is not null and datatype is null we propably have a DateTime field.
            // if datatype would not be null it could also be boolean, as far as i tested it ~Kevin F 13.10.2023
            if cell.StyleIndex <> null && cell.DataType = null then
                try
                    let styleSheet = Stylesheet.get doc 
                    let cellFormat : CellFormat = Stylesheet.CellFormat.getAt (int cell.StyleIndex.InnerText) styleSheet
                    if cellFormat <> null then
                        // if numberformatid is between 14 and 18 it is standard date time format.
                        // custom formats are given in the range of 164 to 180, all none default date time formats fall in there.
                        let dateTimeFormats = [14..22]@[164 .. 180] |> List.map (uint32 >> UInt32Value)
                        if List.contains cellFormat.NumberFormatId dateTimeFormats then 
                            DataType.Date
                        else 
                            DataType.Number
                    else
                        DataType.Number
                with
                | _ -> DataType.Number
            else 
                let cellValues = cell.DataType.Value
                match cellValues with
                | CellValues.Number -> DataType.Number
                | CellValues.Boolean -> DataType.Boolean
                | CellValues.Date -> DataType.Date
                | CellValues.Error -> DataType.Empty
                | CellValues.InlineString
                | CellValues.SharedString
                | CellValues.String -> DataType.String
                | _ -> DataType.Number


    type FsCell with

        /// <summary>
        /// Creates an FsCell on the basis of an XlsxCell. Uses a SharedStringTable if present to get the XlsxCell's value.
        /// </summary>   
        static member ofXlsxCell (doc : Packaging.SpreadsheetDocument) (xlsxCell : Cell) =
            let sst = Spreadsheet.tryGetSharedStringTable doc
            let mutable v =  Cell.getValue sst xlsxCell
            let setValue x = v <- x
            let col, row = xlsxCell.CellReference.Value |> CellReference.toIndices
            let dt = 
                try DataType.ofXlsXCell doc xlsxCell
                with _ -> DataType.Number // default is number 
            match dt with
            | Date ->
                try 
                    // datetime is written as float counting days since 1900. 
                    // We use the .NET helper because we really do not want to deal with datetime issues.
                    setValue <| System.DateTime.FromOADate(float v).ToString()
                with 
                    | _ -> ()
            | Boolean ->
                // boolean is written as int/float either 0 or null
                match v with 
                | "1" -> setValue "true" 
                | "0" -> setValue "false" 
                | _ -> ()
            | _ -> ()
            FsCell.createWithDataType dt (int row) (int col) v

        static member toXlsxCell (doc : Packaging.SpreadsheetDocument) (cell : FsCell) =
            Cell.fromValueWithDataType doc (uint32 cell.ColumnNumber) (uint32 cell.RowNumber) cell.Value cell.DataType

    type FsTable with

        /// <summary>
        /// Returns the FsTable with given FsCellsCollection in the form of an XlsxTable.
        /// </summary>
        member self.ToXlsxTable(cells : FsCellsCollection) = 
            self.RescanFieldNames cells
            let columns =
                self.GetFieldNames(cells)
                |> Seq.map (fun kv -> 
                    Table.TableColumn.create (1 + kv.Value.Index |> uint) kv.Value.Name
                )
            Table.create self.Name (StringValue(self.RangeAddress.Range)) columns

        /// <summary>
        /// Returns an FsTable with given FsCellsCollection in the form of an XlsxTable.
        /// </summary>
        static member toXlsxTable cellsCollection (table : FsTable) =
            table.ToXlsxTable(cellsCollection)

        ///// Creates an FsTable on the basis of an XlsxTable.
        //new(table : Spreadsheet.Table) =        // not permitted :(
            //FsTable(table)

        /// <summary>
        /// Takes an XlsxTable and returns an FsTable.
        /// </summary>
        static member fromXlsxTable table = 
            let topLeftBoundary, bottomRightBoundary = Table.getArea table |> Table.Area.toBoundaries
            let ra = FsRangeAddress(FsAddress(topLeftBoundary), FsAddress(bottomRightBoundary))
            let totalsRowShown = if table.TotalsRowShown = null then false else table.TotalsRowShown.Value
            FsTable(table.DisplayName, ra, totalsRowShown, true)

        /// <summary>
        /// Returns the FsWorksheet associated with the FsTable in a given FsWorkbook.
        /// </summary>
        member self.GetWorksheetOfTable(workbook : FsWorkbook) =
            workbook.GetWorksheets() 
            |> Seq.find (
                fun s -> 
                    s.Tables 
                    |> Seq.exists (fun t -> t.Name = self.Name)
            )

        /// <summary>
        /// Returns the FsWorksheet associated with a given FsTable in an FsWorkbook.
        /// </summary>
        static member getWorksheetOfTable workbook (table : FsTable) =
            table.GetWorksheetOfTable workbook


    type FsWorksheet with

        /// <summary>
        /// Returns the FsWorksheet in the form of an XlsxSpreadsheet.
        /// </summary>
        member self.ToXlsxWorksheet(doc) =
            self.RescanRows()
            let sheet = Worksheet.empty()
            let sheetData =
                let sd = SheetData.empty()
                self.SortRows()
                for row in self.Rows do
                    let cells = row.Cells |> Seq.toList
                    if not cells.IsEmpty then
                        let min,max =
                            cells
                            |> List.map (fun cell -> uint32 cell.ColumnNumber) 
                            |> fun l -> List.min l, List.max l
                        let cells = 
                            cells
                            |> List.map (fun cell ->
                                Cell.fromValueWithDataType doc (uint32 cell.ColumnNumber) (uint32 cell.RowNumber) (cell.Value) (cell.DataType)
                            )
                        let row = Row.create (uint32 row.Index) (Row.Spans.fromBoundaries min max) cells
                        SheetData.appendRow row sd |> ignore
                sd
            Worksheet.setSheetData sheetData sheet

        /// <summary>
        /// Returns an FsWorksheet in the form of an XlsxSpreadsheet.
        /// </summary>
        static member toXlsxWorksheet (fsWorksheet : FsWorksheet, doc) = 
            fsWorksheet.ToXlsxWorksheet(doc)

        /// <summary>
        /// Appends the FsTables of this FsWorksheet to a given OpenXmlWorksheetPart in an XlsxWorkbookPart.
        /// </summary>
        member self.AppendTablesToWorksheetPart(xlsxlWorkbookPart : DocumentFormat.OpenXml.Packaging.WorkbookPart, xlsxWorksheetPart : DocumentFormat.OpenXml.Packaging.WorksheetPart) =
            self.Tables
            |> Seq.iter (fun t ->
                let table = t.ToXlsxTable(self.CellCollection)
                Table.addTable xlsxlWorkbookPart xlsxWorksheetPart table |> ignore
            )

        /// <summary>
        /// Appends the FsTables of an FsWorksheet to a given OpenXmlWorksheetPart in an XlsxWorkbookPart.
        /// </summary>
        static member appendTablesToWorksheetPart xlsxWorkbookPart xlsxWorksheetPart (fsWorksheet : FsWorksheet) =
            fsWorksheet.AppendTablesToWorksheetPart(xlsxWorkbookPart, xlsxWorksheetPart)


    type FsWorkbook with
        
        /// <summary>
        /// Creates an FsWorkbook from a given SpreadsheetDocument.
        /// </summary>
        static member fromSpreadsheetDocument (doc : Packaging.SpreadsheetDocument) =
            let sst = Spreadsheet.tryGetSharedStringTable doc
            let xlsxWorkbookPart = Spreadsheet.getWorkbookPart doc        
            let xlsxSheets = 
                try
                    let xlsxWorkbook = Workbook.get xlsxWorkbookPart
                    Sheet.Sheets.get xlsxWorkbook
                    |> Sheet.Sheets.getSheets
                with 
                | _ -> []
            let xlsxWorksheetParts = 
                xlsxSheets
                |> Seq.map (
                    fun s -> 
                        let sid = Sheet.getID s
                        sid, Worksheet.WorksheetPart.getByID sid xlsxWorkbookPart
                )
            let xlsxTables = 
                xlsxWorksheetParts 
                |> Seq.map (fun (sid, wsp) -> sid, Worksheet.WorksheetPart.getTables wsp)

            let sheets =
                xlsxSheets
                |> Seq.map (
                    fun xlsxSheet ->
                        let sheetIndex = Sheet.getSheetIndex xlsxSheet //unused?
                        let sheetId = Sheet.getID xlsxSheet
                        let xlsxCells = 
                            Spreadsheet.getCellsBySheetID sheetId doc
                            |> Seq.map (FsCell.ofXlsxCell doc)
                        let assocXlsxTables = 
                            xlsxTables 
                            |> Seq.tryPick (fun (sid,ts) -> if sid = sheetId then Some ts else None)
                        let fsTables =
                            match assocXlsxTables with
                            | Some ts -> ts |> Seq.map FsTable.fromXlsxTable |> List.ofSeq
                            | None -> []
                        let fsWorksheet = FsWorksheet(xlsxSheet.Name)
                        fsWorksheet
                        |> FsWorksheet.addCells xlsxCells
                        |> FsWorksheet.addTables fsTables
                )

            sheets
            |> Seq.fold (
                fun wb sheet -> 
                    sheet.RescanRows()      // we need this to have all FsRows present in the final FsWorksheet
                    FsWorkbook.addWorksheet sheet wb
            ) (new FsWorkbook())

        /// <summary>
        /// Creates an FsWorkbook from a given Packaging.Package xlsx package.
        /// </summary>
        static member fromPackage(package:Packaging.Package) =
            let doc = Packaging.SpreadsheetDocument.Open(package)
            FsWorkbook.fromSpreadsheetDocument doc

        /// <summary>
        /// Creates an FsWorkbook from a given Stream to an XlsxFile.
        /// </summary>
        static member fromXlsxStream (stream : Stream) =
            let doc = Spreadsheet.fromStream stream false
            FsWorkbook.fromSpreadsheetDocument doc

        /// <summary>
        /// Creates an FsWorkbook from a given Stream to an XlsxFile.
        /// </summary>
        static member fromBytes (bytes : byte []) =
            let stream = new MemoryStream(bytes)
            FsWorkbook.fromXlsxStream stream

        /// <summary>
        /// Takes the path to an Xlsx file and returns the FsWorkbook based on its content.
        /// </summary>
        static member fromXlsxFile (filePath : string) =
            let sr = new StreamReader(filePath)
            let wb = FsWorkbook.fromXlsxStream sr.BaseStream
            sr.Close()
            wb

        member self.ToEmptySpreadsheet(doc : Packaging.SpreadsheetDocument) =
            
            let workbookPart = Spreadsheet.initWorkbookPart doc

            for worksheet in self.GetWorksheets() do
                let worksheetPart = 
                    WorkbookPart.appendWorksheet worksheet.Name (worksheet.ToXlsxWorksheet(doc)) workbookPart
                    |> WorkbookPart.getOrInitWorksheetPartByName worksheet.Name
               
                worksheet.AppendTablesToWorksheetPart(workbookPart,worksheetPart)

        /// <summary>
        /// Writes the FsWorkbook into a given MemoryStream.
        /// </summary>
        member self.ToStream(stream : MemoryStream) = 
            let doc = Spreadsheet.initEmptyOnStream stream 

            self.ToEmptySpreadsheet(doc)
                //Worksheet.setSheetData sheetData sheet |> ignore
                //WorkbookPart.appendWorksheet worksheet.Name sheet workbookPart |> ignore

            Spreadsheet.close doc

        /// <summary>
        /// Writes an FsWorkbook into a given MemoryStream.
        /// </summary>
        static member toStream stream (workbook : FsWorkbook) =
            workbook.ToStream stream

        /// <summary>
        /// Returns the FsWorkbook in the form of a byte array.
        /// </summary>
        member self.ToBytes() =
            use memoryStream = new MemoryStream()
            self.ToStream(memoryStream)
            memoryStream.ToArray()

        /// <summary>
        /// Returns an FsWorkbook in the form of a byte array.
        /// </summary>
        static member toBytes (workbook: FsWorkbook) =
            workbook.ToBytes()

        /// <summary>
        /// Writes the FsWorkbook into a binary file at the given path.
        /// </summary>
        member self.ToFile(path) =
            self.ToBytes()
            |> fun bytes -> File.WriteAllBytes (path, bytes)

        /// <summary>
        /// Writes an FsWorkbook into a binary file at the given path.
        /// </summary>
        static member toFile path (workbook : FsWorkbook) =
            workbook.ToFile(path)


type Writer =

    /// <summary>
    /// Writes an FsWorkbook into a given MemoryStream.
    /// </summary>
    static member toStream(stream : MemoryStream, workbook : FsWorkbook) =
        workbook.ToStream(stream)

    
    /// <summary>
    /// Returns an FsWorkbook in the form of a byte array.
    /// </summary>
    static member toBytes(workbook: FsWorkbook) =
        workbook.ToBytes()


    /// <summary>
    /// Writes an FsWorkbook into a binary file at the given path.
    /// </summary>
    static member toFile(path,workbook: FsWorkbook) =
        workbook.ToFile(path)