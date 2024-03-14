module Table.Tests

open FsSpreadsheet
open FsSpreadsheet.Net
open DocumentFormat.OpenXml
open TestingUtils


let transformTable =
    testList "transformTable" [
        testCase "handleNullFields" (fun () ->
            let table = FsSpreadsheet.Net.Table.create "TestTable" (StringValue ("A1:D4")) []
            Expect.isTrue (table.TotalsRowShown = null) "Check that field of interest is None"
            FsTable.fromXlsxTable table |> ignore
            )
        testCase "InfersTableColumnsFromRange" (fun () ->
        
            // --- Prepare Table ---
            let wb = new FsWorkbook()

            let ws = wb.InitWorksheet("New Worksheet")

            let columnNames = [|"My Column 1";"My Column 2"|]

            ws.AddCell(FsCell(columnNames.[0],address=FsAddress("B1"))) |> ignore
            ws.AddCell(FsCell(columnNames.[1],address=FsAddress("C1"))) |> ignore
            let t = FsTable("My_New_Table", FsRangeAddress("B1:C2"))

            ws.AddTable(t) |> ignore

            // --- Function of interest ---

            let bytes = wb.ToXlsxBytes()

            // --- Get Tables ---

            let ms = new System.IO.MemoryStream(bytes)

            let doc = Spreadsheet.fromStream ms false
            let xlsxWorkbookPart = Spreadsheet.getWorkbookPart doc       

            let tables = 
                Workbook.get xlsxWorkbookPart
                |> Sheet.Sheets.get
                |> Sheet.Sheets.getSheets
                |> Seq.collect (
                    fun s -> 
                        let sid = Sheet.getID s
                        Worksheet.WorksheetPart.getByID sid xlsxWorkbookPart
                        |> Worksheet.WorksheetPart.getTables
                )


            // --- Checks ---
            Expect.equal (Seq.length tables) 1 "Check that there is one table"

            let table = Seq.head tables
            let tableColumnNames = 
                table.TableColumns
                |> Table.TableColumns.getTableColumns
                |> Seq.map (fun tc -> Table.TableColumn.getName tc)
            
            Expect.sequenceEqual tableColumnNames columnNames "Check that column names match"            
        
        
        )
    
    ]
    

let main =
    testList "FsTable" [
        transformTable
    ]
