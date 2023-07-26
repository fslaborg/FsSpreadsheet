namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open FsSpreadsheet
open FsSpreadsheet.ExcelIO
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
        static member ofXlsxCellValues (cellValues : CellValues) =
            match cellValues with
            | CellValues.Number -> DataType.Number
            | CellValues.Boolean -> DataType.Boolean
            | CellValues.Date -> DataType.Date
            | CellValues.Error -> DataType.Empty
            | CellValues.InlineString
            | CellValues.SharedString
            | CellValues.String 
            | _ -> DataType.String


    type FsCell with
        //member self.ofXlsxCell (sst : Spreadsheet.SharedStringTable option) (xlsxCell:Spreadsheet.Cell) =
        //    let v =  Cell.getValue sst xlsxCell
        //    let row,col = xlsxCell.CellReference.Value |> CellReference.toIndices
        //    FsCell.create (int row) (int col) v

        /// <summary>
        /// Creates an FsCell on the basis of an XlsxCell. Uses a SharedStringTable if present to get the XlsxCell's value.
        /// </summary>
        static member ofXlsxCell (sst : SharedStringTable option) (xlsxCell : Cell) =
            let v =  Cell.getValue sst xlsxCell
            let col, row = xlsxCell.CellReference.Value |> CellReference.toIndices
            let dt = 
                try DataType.ofXlsxCellValues xlsxCell.DataType.Value
                with _ -> DataType.Empty
            FsCell.createWithDataType dt (int row) (int col) v


    type FsTable with

        /// <summary>
        /// Returns the FsTable with given FsCellsCollection in the form of an XlsxTable.
        /// </summary>
        member self.ToXlsxTable(cells : FsCellsCollection) = 

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
            let showHeaderRow = if table.HeaderRowCount = null then false else table.HeaderRowCount.Value = 1u
            FsTable(table.Name, ra, totalsRowShown, showHeaderRow)

        /// <summary>
        /// Returns the FsWorksheet associated with the FsTable in a given FsWorkbook.
        /// </summary>
        member self.GetWorksheetOfTable(workbook : FsWorkbook) =
            workbook.GetWorksheets() 
            |> List.find (
                fun s -> 
                    s.Tables 
                    |> List.exists (fun t -> t.Name = self.Name)
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
        member self.ToXlsxWorksheet() =
            self.RescanRows()
            let sheet = Worksheet.empty()
            let sheetData =
                let sd = SheetData.empty()
                self.SortRows()
                self.Rows
                |> List.iter (fun row -> 
                    let cells = row.Cells |> Seq.toList
                    if not cells.IsEmpty then
                        let min,max =
                            cells
                            |> List.map (fun cell -> uint32 cell.ColumnNumber) 
                            |> fun l -> List.min l, List.max l
                        let cells = 
                            cells
                            |> List.map (fun cell ->
                                Cell.fromValueWithDataType None (uint32 cell.ColumnNumber) (uint32 cell.RowNumber) (cell.Value) (cell.DataType)
                            )
                        let row = Row.create (uint32 row.Index) (Row.Spans.fromBoundaries min max) cells
                        SheetData.appendRow row sd |> ignore
                ) 
                sd
            Worksheet.setSheetData sheetData sheet

        /// <summary>
        /// Returns an FsWorksheet in the form of an XlsxSpreadsheet.
        /// </summary>
        static member toXlsxWorksheet (fsWorksheet : FsWorksheet) = 
            fsWorksheet.ToXlsxWorksheet()

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
        /// Creates an FsWorkbook from a given Stream to an XlsxFile.
        /// </summary>
        // TO DO: Ask HLW/TM: is this REALLY the way to go? This is not a constructor! (though it tries to be one)
        member self.FromXlsxStream (stream : Stream) =
            let doc = Spreadsheet.fromStream stream false
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
                        let sheetIndex = Sheet.getSheetIndex xlsxSheet
                        let sheetId = Sheet.getID xlsxSheet
                        let xlsxCells = 
                            Spreadsheet.getCellsBySheetID sheetId doc
                            |> Seq.map (FsCell.ofXlsxCell sst)
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
        /// Creates an FsWorkbook from a given Stream to an XlsxFile.
        /// </summary>
        static member fromXlsxStream (stream : Stream) =
            (new FsWorkbook()).FromXlsxStream stream

        /// <summary>
        /// Creates an FsWorkbook from a given Stream to an XlsxFile.
        /// </summary>
        static member fromBytes (bytes : byte []) =
            let stream = new MemoryStream(bytes)
            (new FsWorkbook()).FromXlsxStream stream

        /// <summary>
        /// Takes the path to an Xlsx file and returns the FsWorkbook based on its content.
        /// </summary>
        static member fromXlsxFile (filePath : string) =
            let sr = new StreamReader(filePath)
            let wb = FsWorkbook.fromXlsxStream sr.BaseStream
            sr.Close()
            wb

        /// <summary>
        /// Writes the FsWorkbook into a given MemoryStream.
        /// </summary>
        member self.ToStream(stream : MemoryStream) = 
            let doc = Spreadsheet.initEmptyOnStream stream 

            let workbookPart = Spreadsheet.initWorkbookPart doc

            self.GetWorksheets()
            |> List.iter (fun worksheet ->

                let worksheetPart = 
                    WorkbookPart.appendWorksheet worksheet.Name (worksheet.ToXlsxWorksheet()) workbookPart
                    |> WorkbookPart.getOrInitWorksheetPartByName worksheet.Name
               
                worksheet.AppendTablesToWorksheetPart(workbookPart,worksheetPart)
                //Worksheet.setSheetData sheetData sheet |> ignore
                //WorkbookPart.appendWorksheet worksheet.Name sheet workbookPart |> ignore
            )

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