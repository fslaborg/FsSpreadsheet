namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Spreadsheet
open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open System.IO

/// Classes that extend the core FsSpreadsheet library with IO functionalities.
[<AutoOpen>]
module FsExtensions =


    type DataType with

        /// <summary>Converts a given CellValues to the respective DataType.</summary>
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

        /// <summary>Creates an FsCell on the basis of an XlsxCell. Uses a SharedStringTable if present to get the XlsxCell's value.</summary>
        static member ofXlsxCell (sst : SharedStringTable option) (xlsxCell : Cell) =
            let v =  Cell.getValue sst xlsxCell
            let col, row = xlsxCell.CellReference.Value |> CellReference.toIndices
            let dt = DataType.ofXlsxCellValues xlsxCell.DataType
            FsCell.createWithDataType dt (int row) (int col) v


    type FsTable with

        /// Returns the FsTable with given FsCellsCollection in the form of an XlsxTable.
        member self.ToXlsxTable(cells : FsCellsCollection) = 

            let columns =
                self.FieldNames(cells)
                |> Seq.map (fun kv -> 
                    Table.TableColumn.create (1 + kv.Value.Index |> uint) kv.Value.Name
                )
            Table.create self.Name (StringValue(self.RangeAddress.Range)) columns

        /// Returns an FsTable with given FsCellsCollection in the form of an XlsxTable.
        static member toXlsxTable cellsCollection (table : FsTable) =
            table.ToXlsxTable(cellsCollection)

        ///// Creates an FsTable on the basis of an XlsxTable.
        //new(table : Spreadsheet.Table) =        // not permitted :(
            //FsTable(table)

        /// Takes an XlsxTable and returns an FsTable.
        static member fromXlsxTable table = 
            let topLeftBoundary, bottomRightBoundary = Table.getArea table |> Table.Area.toBoundaries
            let ra = FsRangeAddress(FsAddress(topLeftBoundary), FsAddress(bottomRightBoundary))
            FsTable(table.Name, ra, table.TotalsRowShown, true)


    type FsWorksheet with

        /// Returns the FsWorksheet in the form of an XlsxSpreadsheet.
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

        /// Returns an FsWorksheet in the form of an XlsxSpreadsheet.
        static member toXlsxWorksheet (fsWorksheet : FsWorksheet) = 
            fsWorksheet.ToXlsxWorksheet()

        /// Appends the FsTables of this FsWorksheet to a given OpenXmlWorksheetPart in an XlsxWorkbookPart.
        member self.AppendTablesToWorksheetPart(xlsxlWorkbookPart : DocumentFormat.OpenXml.Packaging.WorkbookPart, xlsxWorksheetPart : DocumentFormat.OpenXml.Packaging.WorksheetPart) =
            self.Tables
            |> Seq.iter (fun t ->
                let table = t.ToXlsxTable(self.CellCollection)
                Table.addTable xlsxlWorkbookPart xlsxWorksheetPart table |> ignore
            )

        /// Appends the FsTables of an FsWorksheet to a given OpenXmlWorksheetPart in an XlsxWorkbookPart.
        static member appendTablesToWorksheetPart xlsxWorkbookPart xlsxWorksheetPart (fsWorksheet : FsWorksheet) =
            fsWorksheet.AppendTablesToWorksheetPart(xlsxWorkbookPart, xlsxWorksheetPart)


    type FsWorkbook with
        
        /// <summary>Creates an FsWorkbook from a given Stream to an XlsxFile.</summary>
        // TO DO: Ask HLW/TM: is this REALLY the way to go? This is not a constructor! (though it tries to be one)
        member self.FromXlsxStream (stream : Stream) =
            let doc = Spreadsheet.fromStream stream false
            let sst = Spreadsheet.tryGetSharedStringTable doc
            let xlsxWorkbookPart = Spreadsheet.getWorkbookPart doc
            let xlsxWorkbook = Workbook.get xlsxWorkbookPart
            let xlsxSheets = 
                Sheet.Sheets.get xlsxWorkbook
                |> Sheet.Sheets.getSheets
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
                            Spreadsheet.getCellsBySheetIndex sheetIndex doc
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
            |> Seq.fold (fun wb sheet -> FsWorkbook.addWorksheet sheet wb) (new FsWorkbook())

        /// <summary>Creates an FsWorkbook from a given Stream to an XlsxFile.</summary>
        static member fromXlsxStream (stream : Stream) =
            (new FsWorkbook()).FromXlsxStream stream
            
        //member self.FromFile (path:string) =
        //    use sr = new StreamReader(path)
        //    self.FromXlsxStream sr.

        /// Writes the FsWorkbook into a given MemoryStream.
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

        /// Writes an FsWorkbook into a given MemoryStream.
        static member toStream stream (workbook : FsWorkbook) =
            workbook.ToStream stream

        /// Returns the FsWorkbook in the form of a byte array.
        member self.ToBytes() =
            use memoryStream = new MemoryStream()
            self.ToStream(memoryStream)
            memoryStream.ToArray()

        /// Returns an FsWorkbook in the form of a byte array.
        static member toBytes (workbook: FsWorkbook) =
            workbook.ToBytes()

        /// Writes the FsWorkbook into a binary file at the given path.
        member self.ToFile(path) =
            self.ToBytes()
            |> fun bytes -> File.WriteAllBytes (path, bytes)

        /// Writes an FsWorkbook into a binary file at the given path.
        static member toFile path (workbook : FsWorkbook) =
            workbook.ToFile(path)


type Writer =
    
    /// Writes an FsWorkbook into a given MemoryStream.
    static member toStream(stream : MemoryStream, workbook : FsWorkbook) =
        workbook.ToStream(stream)

    /// Returns an FsWorkbook in the form of a byte array.
    static member toBytes(workbook: FsWorkbook) =
        workbook.ToBytes()

    /// Writes an FsWorkbook into a binary file at the given path.
    static member toFile(path,workbook: FsWorkbook) =
        workbook.ToFile(path)