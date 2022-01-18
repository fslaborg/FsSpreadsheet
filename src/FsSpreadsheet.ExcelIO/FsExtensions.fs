namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml
open FsSpreadsheet

[<AutoOpen>]
module FsExtensions =

    type FsTable with

        member self.ToExcelTable (cells : FsCellsCollection) = 

            let columns =
                self.FieldNames(cells)
                |> Seq.map (fun kv -> 
                    Table.TableColumn.create (1 + kv.Value.Index |> uint) kv.Value.Name
                )
            Table.create self.Name (StringValue(self.RangeAddress.Range)) columns

    type FsWorksheet with

        member self.ToExcelWorksheet() =
            self.RescanRows()
            let sheet = Worksheet.empty()
            let sheetData =                 
                let sd = SheetData.empty()
                self.SortRows()
                self.GetRows()
                |> List.iter (fun row -> 
                    let cells = row.Cells(self.CellCollection) |> Seq.toList
                    let min,max =                        
                        cells
                        |> List.map (fun cell -> uint32 cell.WorksheetColumn) 
                        |> fun l -> List.min l, List.max l
                    let cells = 
                        cells
                        |> List.map (fun cell ->
                            Cell.fromValue None (uint32 cell.WorksheetColumn) (uint32 cell.WorksheetRow) (cell.Value)
                        )
                    let row = Row.create (uint32 row.Index) (Row.Spans.fromBoundaries min max) cells
                    SheetData.appendRow row sd |> ignore
                ) 
                sd
            Worksheet.setSheetData sheetData sheet


        member self.AppendTablesToWorksheetPart(worksheetPart : DocumentFormat.OpenXml.Packaging.WorksheetPart) =
            self.GetTables()
            |> Seq.iter (fun t ->
                let table = t.ToExcelTable(self.CellCollection)
                Table.addTable worksheetPart table |> ignore
            )


    type FsWorkbook with

        member self.ToStream(stream : System.IO.MemoryStream) = 
            let doc = Spreadsheet.initEmptyOnStream stream 

            let workbookPart = Spreadsheet.initWorkbookPart doc

            self.GetWorksheets()
            |> List.iter (fun worksheet ->

                let worksheetPart = 
                    WorkbookPart.appendWorksheet worksheet.Name (worksheet.ToExcelWorksheet()) workbookPart
                    |> WorkbookPart.getOrInitWorksheetPartByName worksheet.Name
               
                worksheet.AppendTablesToWorksheetPart(worksheetPart)
                //Worksheet.setSheetData sheetData sheet |> ignore
                //WorkbookPart.appendWorksheet worksheet.Name sheet workbookPart |> ignore
            )

            Spreadsheet.close doc

        member self.ToBytes() =
            use memoryStream = new System.IO.MemoryStream()
            self.ToStream(memoryStream)
            memoryStream.ToArray()

        member self.ToFile(path) =
            self.ToBytes()
            |> fun bytes -> System.IO.File.WriteAllBytes (path, bytes)

        static member toStream(stream : System.IO.MemoryStream,workbook : FsWorkbook) =
            workbook.ToStream(stream)

        static member toBytes(workbook: FsWorkbook) =
            workbook.ToBytes()

        static member toFile(path,workbook: FsWorkbook) =
            workbook.ToFile(path)
