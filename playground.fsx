#r "nuget: DocumentFormat.OpenXml"
//#load "src/FsSpreadsheet/FsAddress.fs"
//#load "src/FsSpreadsheet/Cells/FsCell.fs"
//#load "src/FsSpreadsheet/Cells/FsCellsCollection.fs"
//#load "src/FsSpreadsheet/Ranges/FsRangeAddress.fs"
//#load "src/FsSpreadsheet/Ranges/FsRangeBase.fs"
//#load "src/FsSpreadsheet/Ranges/FsRangeRow.fs"
//#load "src/FsSpreadsheet/Ranges/FsRangeColumn.fs"
//#load "src/FsSpreadsheet/Ranges/FsRange.fs"
//#load "src/FsSpreadsheet/Tables/FsTableField.fs"
//#load "src/FsSpreadsheet/Tables/FsTableRow.fs"
//#load "src/FsSpreadsheet/Tables/FsTable.fs"
//#load "src/FsSpreadsheet/FsRow.fs"
//#load "src/FsSpreadsheet/FsWorksheet.fs"
//#load "src/FsSpreadsheet/FsWorkbook.fs"
//#load "src/FsSpreadsheet/DSL/Expression.fs"
//#load "src/FsSpreadsheet/DSL/Types.fs"
//#load "src/FsSpreadsheet/DSL/CellBuilder.fs"
//#load "src/FsSpreadsheet/DSL/RowBuilder.fs"
//#load "src/FsSpreadsheet/DSL/ColumnBuilder.fs"
//#load "src/FsSpreadsheet/DSL/TableBuilder.fs"
//#load "src/FsSpreadsheet/DSL/SheetBuilder.fs"
//#load "src/FsSpreadsheet/DSL/WorkbookBuilder.fs"
//#load "src/FsSpreadsheet/DSL/DSL.fs"
//#load "src/FsSpreadsheet/DSL/Operators.fs"
//#load "src/FsSpreadsheet/DSL/Transform.fs"
//#load "src/FsSpreadsheet.CsvIO/FsExtension.fs"
//#load "src/FsSpreadsheet.ExcelIO/SharedStringTable.fs"
//#load "src/FsSpreadsheet.ExcelIO/Cell.fs"
//#load "src/FsSpreadsheet.ExcelIO/CellData.fs"
//#load "src/FsSpreadsheet.ExcelIO/Row.fs"
//#load "src/FsSpreadsheet.ExcelIO/SheetData.fs"
//#load "src/FsSpreadsheet.ExcelIO/Sheet.fs"
//#load "src/FsSpreadsheet.ExcelIO/Table.fs"
//#load "src/FsSpreadsheet.ExcelIO/WorkSheet.fs"
//#load "src/FsSpreadsheet.ExcelIO/Workbook.fs"
//#load "src/FsSpreadsheet.ExcelIO/Spreadsheet.fs"
//#load "src/FsSpreadsheet.ExcelIO/FsExtensions.fs"
#r "src/FsSpreadsheet/bin/Debug/netstandard2.0/FsSpreadsheet.dll"
#r "src/FsSpreadsheet.CsvIO/bin/Debug/netstandard2.0/FsSpreadsheet.CsvIO.dll"
#r "src/FsSpreadsheet.ExcelIO/bin/Debug/netstandard2.0/FsSpreadsheet.ExcelIO.dll"


open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open FsSpreadsheet.DSL
open DocumentFormat.OpenXml


// ----------------------------------------------

//let excelFilePath = @"C:\Users\olive\OneDrive\CSB-Stuff\testFiles\testExcel5.xlsx"
let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel5.xlsx"

let dslTree = 
    workbook {
        sheet "MySheet" {
            row {
                cell {1}
                cell {2}
                cell {3}
            }
            row {
                4
                5
                6
            }
        }
    }

let fsWorkbook = dslTree.Value.Parse()
fsWorkbook.ToFile(excelFilePath)

let doc = Spreadsheet.fromFile excelFilePath false
let sst = Spreadsheet.tryGetSharedStringTable doc
let wbp = Spreadsheet.getWorkbookPart doc
let wb = Workbook.get wbp
let shts = Sheet.Sheets.get wb
let shtsN = Sheet.Sheets.getSheets shts |> Array.ofSeq
shtsN.Length
shtsN[0].Id.Value
let wspN =
    shtsN
    |> Array.map (
        fun s -> Worksheet.WorksheetPart.getByID s.Id.Value wbp
    )
let tblNM = wspN |> Array.map (Worksheet.WorksheetPart.getTables >> Array.ofSeq)
let wsN = wspN |> Array.map Worksheet.get
let sdN = wsN |> Array.map Worksheet.getSheetData
let rNM = sdN |> Array.map (SheetData.getRows >> Array.ofSeq)
let cNMO =
    rNM                                             // Level 1: Worksheets (N)
    |> Array.map (                                  // Level 2: Rows (M)
        Array.map (Row.toCellSeq >> Array.ofSeq)    // Level 3: Cells (O)
    )
let testCell = cNMO[0].[0].[0]
let cv = Cell.getValue sst testCell
let dt = Cell.getType testCell |> Cell.cellValuesToDataType
let cr = Cell.getReference testCell
let fa = FsAddress cr
FsCell(cv, dt, fa)
let fswb = new FsWorkbook()
shtsN[0]
let name = Sheet.getName shtsN[0]
let tblN = tblNM[0]
//tblN[0].TableDefinitionPart
//Table.
let fsTblN = tblN |> Array.map FsTable.fromXlsxTable
let fsCcN = FsCellsCollection()
cNMO[0]
|> Array.iteri (
    fun iR r ->
        r
        |> Array.iteri (
            fun iC c ->
                let cv = Cell.getValue sst c
                let dt = Cell.getType c |> Cell.cellValuesToDataType
                let fa = Cell.getReference c |> FsAddress
                let fsc = FsCell(cv, dt, fa)
                fsCcN.Add(iR, iC, fsc)
        )
)
let sdName = shtsN[0].Name.Value
let fsRs = 
    rNM[0]
    |> Array.map (
        fun r ->
            let fi, li = Row.Spans.toBoundaries r.Spans
            let ri = Row.getIndex r
            let fa = FsAddress(int ri, int fi)
            let la = FsAddress(int ri, int li)
            let fsra = FsRangeAddress(fa, la)
            let fsCci = fsCcN.GetCellsInRow (int ri)
            FsRow(fsra, fsCcN, box 0)
    )
let fsws = FsWorksheet(sdName, )
let fswsN = 
    shtsN
    |> Array.mapi (
        fun i sht -> 
            let name = Sheet.getName sht
            let tblN = tblNM[i]
            let fsTblN = tblN |> Array.map FsTable.fromXlsxTable
            let fsCcN
            FsWorksheet(name, fsRs = , )
    )

let xlsxWorksheet = 0

let toFsWorkbook spreadsheetDoc =
    let sst = Spreadsheet.tryGetSharedStringTable spreadsheetDoc
    let wbp = Spreadsheet.getWorkbookPart spreadsheetDoc
    let wb = Workbook.get wbp
    let shts = Sheet.Sheets.get wb
    let shtsN = Sheet.Sheets.getSheets shts |> Array.ofSeq      // N, M, O = multiples of the elements, e.g. here: multiple Sheet elements
    let wspN =
        shtsN
        |> Array.map (
            fun s -> Worksheet.WorksheetPart.getByID s.Id.Value wbp
        )
    let tblNM = wspN |> Array.map (Worksheet.WorksheetPart.getTables >> Array.ofSeq)
    let wsN = wspN |> Array.map Worksheet.get
    let sdN = wsN |> Array.map Worksheet.getSheetData
    let rNM = sdN |> Array.map (SheetData.getRows >> Array.ofSeq)
    let cNMO =
        rNM                                             // Level 1: Worksheets (N)
        |> Array.map (                                  // Level 2: Rows (M)
            Array.map (Row.toCellSeq >> Array.ofSeq)    // Level 3: Cells (O)
        )
    let fswb = new FsWorkbook()
    let fswsN = 
        shtsN
        |> Array.mapi (
            fun i sht -> 
                let name = Sheet.getName sht
                let tblN = tblNM[i]
                let fsTblN = tblN |> Array.map FsTable.fromXlsxTable
                let fsCcN
                FsWorksheet(name, fsRs = , )
        )
    0


let fromXlsxWorksheet xlsxWorksheet =
    let sd = Worksheet.getSheetData xlsxWorksheet
    xlsxWorksheet.SheetProperties.LocalName
    FsWorksheet()
    0

doc.Close()