namespace FsSpreadsheet.ExcelIO

open FsSpreadsheet

[<AutoOpen>]
module FsExtensions =

    type FsWorkbook with

        member self.ToStream(stream : System.IO.MemoryStream) = 
            let doc = Spreadsheet.initEmptyOnStream stream 

            let workbookPart = Spreadsheet.initWorkbookPart doc

            self.GetWorksheets()
            |> List.iter (fun worksheet ->
                let sheetData =                 
                    let sd = SheetData.empty()
                    worksheet.SortRows()
                    worksheet.GetRows()
                    |> List.iter (fun row -> 
                        let cells = row.Cells(worksheet.CellCollection) |> Seq.toList
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
                WorkbookPart.appendSheet worksheet.Name sheetData workbookPart |> ignore
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
