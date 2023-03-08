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
open System.IO


// ----------------------------------------------

//let excelFilePath = @"C:\Users\olive\OneDrive\CSB-Stuff\testFiles\testExcel5.xlsx"
//let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel5.xlsx"
//let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel6.xlsx"
let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel6_rewritten.xlsx"

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
tblN[0] |> FsTable.fromXlsxTable
//tblN[0] |> Table.
let fsCcN = FsCellsCollection()
cNMO[0]
|> Array.iteri (
    fun iR r ->
        printfn $"iR: {iR}"
        r
        |> Array.iteri (
            fun iC c ->
                printfn $"iC: {iC}"
                let cv = Cell.getValue sst c
                printfn $"cellRef: {Cell.getReference c}, cellV: {cv}"
                printfn $"CellDataType?: {c.DataType}"
                let dt = 
                    let cvo : Spreadsheet.CellValues option = Cell.tryGetType c
                    if cvo.IsSome then
                        Cell.cellValuesToDataType cvo.Value
                    else DataType.InferCellValue cv |> fst
                    //match cvo with    // errors. WHY!?
                    //| Some ct -> Cell.cellValuesToDataType ct
                    //| None -> DataType.InferCellValue cv |> snd
                printfn "dt got"
                let fa = Cell.getReference c |> FsAddress
                printfn $"Fa: {fa.Address}"
                let fsc = FsCell(cv, dt, fa)
                printfn $"Fsc: Row: {fsc.Address.RowNumber} Col: {fsc.Address.ColumnNumber}"
                fsCcN.Add(iR + 1, iC + 1, fsc)
                ()
        )
)
"A1" |> FsAddress
"B1" |> FsAddress
"C2" |> FsAddress
fsCcN.GetCells() |> Array.ofSeq |> Array.iter (fun fsc -> printfn $"Row: {fsc.Address.RowNumber} Col: {fsc.Address.ColumnNumber} Val: {fsc.Value}")
//let bla = if 1 > 0 then failwith "sheesh"
let sdName = shtsN[0].Name.Value
let fsRs = 
    rNM[0]
    |> Array.map (
        fun r ->
            //let r = rNM[0].[0]
            let fi, li = Row.Spans.toBoundaries r.Spans
            let ri = Row.getIndex r
            printfn $"{ri}"
            let fa = FsAddress(int ri, int fi)
            let la = FsAddress(int ri, int li)
            let fsra = FsRangeAddress(fa, la)
            let fscseq = fsCcN.GetCellsInRow (int ri)
            let fscCr = FsCellsCollection()
            fscseq
            |> Seq.iter (fun fsc -> fscCr.Add(int ri, fsc.Address.ColumnNumber, fsc))
            FsRow(fsra, fscCr, box 0)
    )
fsRs |> Array.map (fun r -> r.Cells |> Seq.toArray)
let fsws = FsWorksheet(sdName, fsRs |> List.ofArray, Array.toList fsTblN, fsCcN)
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
fswb.AddWorksheet fsws
let newExcelPath =
    let fi = FileInfo excelFilePath
    let rawFn = Path.GetFileNameWithoutExtension excelFilePath
    let newFn = rawFn + "_rewritten.xlsx"
    Path.Combine(fi.Directory.FullName, newFn)
FsWorkbook.toFile newExcelPath fswb




let testHeaderCells = [
    FsCell("H1", DataType.String, FsAddress(1,1))
    FsCell("H2", DataType.String, FsAddress(1,2))
    FsCell("H3", DataType.String, FsAddress(1,3))
]
let testCells1 =
    Array.init 3 (
        fun i ->
            FsCell($"{i * 2}", DataType.Number, FsAddress(2,i + 1))
    )
let testColl = FsCellsCollection()
testHeaderCells |> List.iter (fun c -> testColl.Add(c.Address.RowNumber, c.Address.ColumnNumber, c))
testCells1 |> Array.iter (fun c -> testColl.Add(c.Address.RowNumber, c.Address.ColumnNumber, c))
let testRangeAddress = FsRangeAddress(testHeaderCells.Head.Address, testCells1[testCells1.Length - 1].Address)
let testTab = FsTable("lel", testRangeAddress)
testTab |> FsTable.toXlsxTable testColl
FsRow.
let testWs = FsWorksheet("testSheet", )

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