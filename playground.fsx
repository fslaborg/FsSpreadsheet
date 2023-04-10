open System.IO

//Directory.GetCurrentDirectory()
//File.Copy("src/FsSpreadsheet/bin/Debug/netstandard2.0/FsSpreadsheet.dll", "src/FsSpreadsheet/bin/Debug/netstandard2.0/FsSpreadsheet_Copy.dll", true)
//File.Copy("src/FsSpreadsheet.CsvIO/bin/Debug/netstandard2.0/FsSpreadsheet.CsvIO.dll", "src/FsSpreadsheet.CsvIO/bin/Debug/netstandard2.0/FsSpreadsheet.CsvIO_Copy.dll", true)
//File.Copy("src/FsSpreadsheet.ExcelIO/bin/Debug/netstandard2.0/FsSpreadsheet.ExcelIO.dll", "src/FsSpreadsheet.ExcelIO/bin/Debug/netstandard2.0/FsSpreadsheet.ExcelIO_Copy.dll", true)

#r "src/FsSpreadsheet/bin/Debug/netstandard2.0/FsSpreadsheet.dll"
#r "src/FsSpreadsheet.CsvIO/bin/Debug/netstandard2.0/FsSpreadsheet.CsvIO.dll"
#r "src/FsSpreadsheet.ExcelIO/bin/Debug/netstandard2.0/FsSpreadsheet.ExcelIO.dll"

#r "nuget: DocumentFormat.OpenXml"


open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open FsSpreadsheet.DSL
open DocumentFormat.OpenXml


// ----------------------------------------------

// some other bugfixes

//let lolCell = Cell.fromValueWithDataType None 1u 1u "3/3" DataType.String
let lolCell = Cell.create Spreadsheet.CellValues.String "A1" <| Spreadsheet.CellValue("3/3")
//let lolCell = Cell.create Spreadsheet.CellValues.InlineString "A1" <| Spreadsheet.CellValue("3/3")
let ssd = Spreadsheet.init "test" @"C:\Users\revil\Downloads\test.xlsx"
let wbp = Spreadsheet.getWorkbookPart ssd
let wb = Workbook.get wbp
let shts = Sheet.Sheets.get wb
let sht1 = Sheet.Sheets.getFirstSheet shts
let sd1 = Spreadsheet.tryGetSheetBySheetName "test" ssd |> Option.get
SheetData.tryGetCellAt 1u 1u sd1
let rw = Row.empty()
Row.getRowValues None rw |> Seq.toArray
Row.setIndex 1u rw
Row.getIndex rw
Row.appendCell lolCell rw
SheetData.appendRow rw sd1
Sheet.countSheets ssd
ssd.Save()
ssd.Close()

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
let spreadsheet = dslTree.Value.Parse()
spreadsheet.GetWorksheets().Head.CellCollection.GetCells() |> List.ofSeq

let assayTf = @"C:\Users\olive\OneDrive\CSB-Stuff\NFDI\testARC30\assays\aid\isa.assay.xlsx"
let xlSsd = Spreadsheet.fromFile assayTf false
let cs = Spreadsheet.getCellsBySheetIndex 2u xlSsd |> Array.ofSeq
for cell in cs do
    printfn "Ref: %s    DT: %A    StyleIndex: %A" cell.CellReference.Value cell.DataType cell.StyleIndex

let kfsTestFile = @"C:\Users\olive\Downloads\2EXT02_Protein (1).xlsx"
let wb = FsWorkbook.fromXlsxFile kfsTestFile
let ws = wb.GetWorksheetByName "2EXT02_ProteinTest"
ws.CellCollection.TryGetCell(2,1)

let x = FsCell.create 1 1 "Kevin"
let y = x.Copy()
x.Value
y.Value <- "Oliver"
ws.CellCollection.RemoveCellAt(1, 1)
//ws.CellCollection.Add


//type TestClass() =
//    member val Value = 10 with get, set
//    member this.Copy() = 
//        let o = TestClass()
//        o.Value <- this.Value
//        o

//let testObjs = Seq.init 4 (fun i -> TestClass() |> fun x -> x.Value <- i; x)
//let testObjs2 = testObjs |> Seq.map (fun o -> o.Copy())
//testObjs |> Seq.item 0 |> fun x -> x.Value
//testObjs2 |> Seq.item 0 |> fun x -> x.Value
//testObjs2 |> Seq.iter (fun x -> x.Value <- 8)
//testObjs2 |> Array.ofSeq |> Array.map (fun x -> x.Value <- 8; x) |> Array.item 0 |> fun x -> x.Value
//let testObjs3 = testObjs2 |> Array.ofSeq
//testObjs3[0]
//testObjs3 |> Array.iter (fun x -> x.Value <- 9)
//let testObj = TestClass()
//testObj.Value
//testObj.Value <- 9
//let testObjs4 = Array.init 4 (fun i -> TestClass())
//testObjs4[0].Value
//let testObjs5 = testObjs4 |> Array.map (fun o -> o.Copy())
//testObjs5[0].Value
//testObjs5[0].Value <- 5
//testObjs5[0].Value
//let testObjs6 = Seq.init 2 (fun _ -> TestClass())
//testObjs6 |> Seq.item 0 |> fun x -> x.Value
//testObjs6 |> Seq.item 0 |> fun x -> x.Value <- 5
//for i in testObjs6 do i.Value <- 5
//testObjs6 |> Seq.item 0 |> fun x -> x.Value

//let cellSeq = List.init 2 (fun i -> FsCell.create (i + 1) 1 "v")
//let testFCC = FsCellsCollection()
//testFCC.Add cellSeq
//testFCC.GetCells()
//let testCells = testFCC.GetCells() |> Seq.map (fun c -> c.Copy())
//let testFCC2 = FsCellsCollection()
//testFCC2.Add testCells
//testFCC2.GetCells() |> Seq.map (fun c -> c.Value <- "v2")
//testFCC2.GetCells()

//let testFsRangeAddress = FsRangeAddress("C1:C3")
//let testFsRangeColumn = FsRangeColumn testFsRangeAddress
//let testFsTableField = FsTableField("testName", 3, testFsRangeColumn)
//let testFsCellsCollection = FsCellsCollection()
//let testFsCells = [
//    FsCell.createWithDataType DataType.String 1 3 "I am the Header!"
//    FsCell.createWithDataType DataType.String 2 3 "first data cell"
//    FsCell.createWithDataType DataType.String 3 3 "second data cell"
//    FsCell.createWithDataType DataType.String 1 4 "Another Header"
//    FsCell.createWithDataType DataType.String 2 4 "first data cell in B col"
//    FsCell.createWithDataType DataType.String 3 4 "second data cell in B col"
//]
//testFsCellsCollection.Add testFsCells |> ignore

//let headerCell = testFsTableField.HeaderCell(testFsCellsCollection, true)
////testFsTableField.SetName("testName2", testFsCellsCollection, true)
//testFsTableField.DataCells(testFsCellsCollection, true)
//testFsTableField.Index <- 5
//testFsTableField.Column
//match testFsTableField.Column with
//| null -> printfn "null!"
//| _ -> printfn "not null!"


//// test basic stuff

//let mutable value1 = "hallo"
//let mutable value2 = value1
//value2 <- "Welt"

//let dummyFsCellsCollection = FsCellsCollection()
//let dummyFsCellsCollection2 = dummyFsCellsCollection
//let dummyFsCellsCollection3 = FsCellsCollection()
//dummyFsCellsCollection = dummyFsCellsCollection2
//dummyFsCellsCollection = dummyFsCellsCollection3
//System.Object.Equals(dummyFsCellsCollection, dummyFsCellsCollection2)
//System.Object.Equals(dummyFsCellsCollection, dummyFsCellsCollection3)
//dummyFsCellsCollection.Count
//dummyFsCellsCollection2.Count
//dummyFsCellsCollection2.Add(FsCell.createEmpty())
//dummyFsCellsCollection2.Count
//dummyFsCellsCollection.Count    // changes too...
////let dummyFsCellsCollection3 : FsCellsCollection = dummyFsCellsCollection.MemberwiseClone() :?> FsCellsCollection

//let array1 = [|"Hallo"|]
//let array2 = array1
//array2[0] <- "Welt"

//let copy (this : FsCellsCollection) =
//    let newCellsColl = FsCellsCollection()
//    let cells : seq<FsCell> = this.GetCells()
//    newCellsColl.Add cells

//let testCell = FsCell.create 1 1 "Hallo"
//let dummyFsCellsCollection4 = FsCellsCollection().Add testCell
//let dummyFsCellsCollection5 = copy dummyFsCellsCollection4
//dummyFsCellsCollection5.TryGetCell(1, 1) |> fun c -> c.Value.Value <- "Welt"
//dummyFsCellsCollection5.TryGetCell(1, 1) |> fun c -> c.Value
//dummyFsCellsCollection4.TryGetCell(1, 1) |> fun c -> c.Value

////let excelFilePath = @"C:\Users\olive\OneDrive\CSB-Stuff\testFiles\testExcel5.xlsx"
////let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel5.xlsx"
////let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel6.xlsx"
//let excelFilePath = @"C:\Users\olive\OneDrive\CSB-Stuff\testFiles\testExcel6.xlsx"
////let excelFilePath = @"C:\Users\revil\OneDrive\CSB-Stuff\testFiles\testExcel6_rewritten.xlsx"
////let excelFilePath = @"C:\Users\olive\OneDrive\CSB-Stuff\testFiles\testExcel6_rewritten.xlsx"


//// inb4 unit tests

//let unitTestFilePath = @"C:\Repos\CSBiology\FsSpreadsheet\tests\FsSpreadsheet.ExcelIO.Tests\data\testUnit.xlsx"
//let sr = new StreamReader(unitTestFilePath)
//let fsWorkbookFromStream = FsWorkbook.fromXlsxStream sr.BaseStream
//sr.Close()
//let fsWorksheet1FromStream = fsWorkbookFromStream.GetWorksheetByName "StringSheet"
//let fsWorksheet2FromStream = fsWorkbookFromStream.GetWorksheetByName "NumericSheet"
//let fsWorksheet3FromStream = fsWorkbookFromStream.GetWorksheetByName "TableSheet"
//let fsWorksheet4FromStream = fsWorkbookFromStream.GetWorksheetByName "DataTypeSheet"
////let v = (FsWorksheet.getCellAt 1 1 fsWorksheet1FromStream).Value
////let a = (FsWorksheet.getCellAt 1 1 fsWorksheet1FromStream).Address.Address
////let d = (FsWorksheet.getCellAt 1 1 fsWorksheet1FromStream).DataType
//let v = (FsWorksheet.getCellAt 7 3 fsWorksheet2FromStream).Value
//fsWorksheet2FromStream.CellCollection.GetCells() |> Array.ofSeq
//let t = fsWorksheet3FromStream.Tables |> List.tryFind (fun t -> t.Name = "Table2")


//// fix some bugs

//#r "nuget: DocumentFormat.OpenXml"

//open DocumentFormat.OpenXml

//let sdoc = Spreadsheet.fromFile excelFilePath false
////let cell = Cell.fromValueWithDataType None 1u 1u "test" DataType.Number
//let cells = Spreadsheet.getCellsBySheetIndex 1u sdoc |> Array.ofSeq
//cells |> Array.iter (fun (c) -> printfn $"Ref: {c.CellReference.Value}   Value: {c.CellValue.Text}")
//let ssdFox = Packaging.SpreadsheetDocument.Open(excelFilePath, false)
//let wbpFox = ssdFox.WorkbookPart
//let sstpFox = wbpFox.SharedStringTablePart
//let sstFox = sstpFox.SharedStringTable
//let sstFoxInnerText = sstFox.InnerText
//let wsp1Fox = (wbpFox.WorksheetParts |> Array.ofSeq)[0]
//let cbsi1Fox = 
//    wsp1Fox.Worksheet.Descendants<Spreadsheet.Cell>() |> Array.ofSeq
//    |> Array.map (
//        fun c ->
//            if c.DataType <> null && c.DataType.Value = Spreadsheet.CellValues.SharedString then
//                let index = int c.CellValue.InnerText
//                let item = sstFox.Elements<OpenXmlElement>() |> Seq.item index
//                let value = item.InnerText
//                c.CellValue.Text <- value
//                c
//            else
//                c
//    )
//cbsi1Fox |> Array.iter (fun c -> printfn $"Ref: {c.CellReference.Value}   Value: {c.CellValue.Text}") 


//// test principles of the implementation

//let someTestCells = 
//    [
//        [
//            // TO DO: ask: why is ///-text from methods/constructors not shown?!
//            FsCell("H1", DataType.String, FsAddress(1,1))
//            FsCell("H2", DataType.String, FsAddress(1,2))
//            FsCell("H3", DataType.String, FsAddress(1,3))
//        ]
//        List.init 3 (
//            fun i ->
//                FsCell($"{i * 2}", DataType.Number, FsAddress(2,i + 1))
//        )
//    ]
//let testFsCC = FsCellsCollection()
//someTestCells |> List.iter (testFsCC.Add >> ignore)
//let testFsRow = FsRow(2, testFsCC)
//(testFsRow.Cells |> Array.ofSeq).Length
//testFsRow.RangeAddress

//let testFsWs = FsWorksheet("test")
//testFsWs.CellCollection
//testFsWs.Row(1)
//FsRow(1, testFsWs.CellCollection)





//let dslTree = 
//    workbook {
//        sheet "MySheet" {
//            row {
//                cell {1}
//                cell {2}
//                cell {3}
//            }
//            row {
//                4
//                5
//                6
//            }
//        }
//    }

//let fsWorkbook = dslTree.Value.Parse()
//fsWorkbook.ToFile(excelFilePath)

//let doc = Spreadsheet.fromFile excelFilePath false
//let sst = Spreadsheet.tryGetSharedStringTable doc
//let wbp = Spreadsheet.getWorkbookPart doc
//let wb = Workbook.get wbp
//let shts = Sheet.Sheets.get wb
//let shtsN = Sheet.Sheets.getSheets shts |> Array.ofSeq
//shtsN.Length
//shtsN[0].Id.Value
//let wspN =
//    shtsN
//    |> Array.map (
//        fun s -> Worksheet.WorksheetPart.getByID s.Id.Value wbp
//    )
//let tblNM = wspN |> Array.map (Worksheet.WorksheetPart.getTables >> Array.ofSeq)
//let wsN = wspN |> Array.map Worksheet.get
//let sdN = wsN |> Array.map Worksheet.getSheetData
//let rNM = sdN |> Array.map (SheetData.getRows >> Array.ofSeq)
//let cNMO =
//    rNM                                             // Level 1: Worksheets (N)
//    |> Array.map (                                  // Level 2: Rows (M)
//        Array.map (Row.toCellSeq >> Array.ofSeq)    // Level 3: Cells (O)
//    )
//let testCell = cNMO[0].[0].[0]
//let cv = Cell.getValue sst testCell
//let dt = Cell.getType testCell |> Cell.cellValuesToDataType
//let cr = Cell.getReference testCell
//let fa = FsAddress cr
//FsCell(cv, dt, fa)
//let fswb = new FsWorkbook()
//shtsN[0]
//let name = Sheet.getName shtsN[0]
//let tblN = tblNM[0]
////tblN[0].TableDefinitionPart
////Table.
//let fsTblN = tblN |> Array.map FsTable.fromXlsxTable
//tblN[0] |> FsTable.fromXlsxTable
////tblN[0] |> Table.
//let fsCcN = FsCellsCollection()
//cNMO[0]
//|> Array.iteri (
//    fun iR r ->
//        printfn $"iR: {iR}"
//        r
//        |> Array.iteri (
//            fun iC c ->
//                printfn $"iC: {iC}"
//                let cv = Cell.getValue sst c
//                printfn $"cellRef: {Cell.getReference c}, cellV: {cv}"
//                printfn $"CellDataType?: {c.DataType}"
//                let dt = 
//                    let cvo : Spreadsheet.CellValues option = Cell.tryGetType c
//                    if cvo.IsSome then
//                        Cell.cellValuesToDataType cvo.Value
//                    else DataType.InferCellValue cv |> fst
//                    //match cvo with    // errors. WHY!?
//                    //| Some ct -> Cell.cellValuesToDataType ct
//                    //| None -> DataType.InferCellValue cv |> snd
//                printfn "dt got"
//                let fa = Cell.getReference c |> FsAddress
//                printfn $"Fa: {fa.Address}"
//                let fsc = FsCell(cv, dt, fa)
//                printfn $"Fsc: Row: {fsc.Address.RowNumber} Col: {fsc.Address.ColumnNumber}"
//                fsCcN.Add(iR + 1, iC + 1, fsc)
//                ()
//        )
//)
//"A1" |> FsAddress
//"B1" |> FsAddress
//"C2" |> FsAddress
//fsCcN.GetCells() |> Array.ofSeq |> Array.iter (fun fsc -> printfn $"Row: {fsc.Address.RowNumber} Col: {fsc.Address.ColumnNumber} Val: {fsc.Value}")
////let bla = if 1 > 0 then failwith "sheesh"

//let sdName = shtsN[0].Name.Value
//let fsRs = 
//    rNM[0]
//    |> Array.map (
//        fun r ->
//            //let r = rNM[0].[0]
//            let fi, li = Row.Spans.toBoundaries r.Spans
//            let ri = Row.getIndex r
//            printfn $"{ri}"
//            let fa = FsAddress(int ri, int fi)
//            let la = FsAddress(int ri, int li)
//            let fsra = FsRangeAddress(fa, la)
//            let fscseq = fsCcN.GetCellsInRow (int ri)
//            let fscCr = FsCellsCollection()
//            fscseq
//            |> Seq.iter (fun fsc -> fscCr.Add(int ri, fsc.Address.ColumnNumber, fsc) |> ignore)
//            FsRow(fsra, fscCr, box 0)
//    )
//fsRs |> Array.map (fun r -> r.Cells |> Seq.toArray)
//let fsws = FsWorksheet(sdName, fsRs |> List.ofArray, Array.toList fsTblN, fsCcN)
//let fswsN = 
//    shtsN
//    |> Array.mapi (
//        fun i sht -> 
//            let name = Sheet.getName sht
//            let tblN = tblNM[i]
//            let fsTblN = tblN |> Array.map FsTable.fromXlsxTable
//            let fsCcN
//            FsWorksheet(name, fsRs = , )
//    )
//fswb.AddWorksheet fsws
//let newExcelPath =
//    let fi = FileInfo excelFilePath
//    let rawFn = Path.GetFileNameWithoutExtension excelFilePath
//    let newFn = rawFn + "_rewritten.xlsx"
//    Path.Combine(fi.Directory.FullName, newFn)
//FsWorkbook.toFile newExcelPath fswb



//let object = FsCell()
//object.Address
//let expectedAddress = FsAddress(0,0)
//let actualAddress = object.Address
//expectedAddress.ColumnNumber = actualAddress.ColumnNumber

///// Checks if 2 FsAddresses are the equal.
//let compareFsAddress (address1 : FsAddress) (address2 : FsAddress) =
//    address1.Address        = address2.Address      &&
//    address1.ColumnNumber   = address2.ColumnNumber &&
//    address1.RowNumber      = address2.RowNumber    &&
//    address1.FixedColumn    = address2.FixedColumn  &&
//    address1.FixedRow       = address2.FixedRow

//let stringValTest = 255uy
//let resultDtTest, resultStrTest = DataType.InferCellValue stringValTest
//let x = box stringValTest
//match x with
//| :? char -> ""
//| _ -> "fif"
//x.ToString()

//type Test() =
//    member val Prop = 0 with get, set

//let test1 = Test()
//test1.Prop
//test1.Prop <- 3


//let testHeaderCells = [
//    FsCell("H1", DataType.String, FsAddress(1,1))
//    FsCell("H2", DataType.String, FsAddress(1,2))
//    FsCell("H3", DataType.String, FsAddress(1,3))
//]
//let testCells1 =
//    Array.init 3 (
//        fun i ->
//            FsCell($"{i * 2}", DataType.Number, FsAddress(2,i + 1))
//    )
//let testColl = FsCellsCollection()
//testHeaderCells |> List.iter (fun c -> testColl.Add(c.Address.RowNumber, c.Address.ColumnNumber, c) |> ignore)
//testCells1 |> Array.iter (fun c -> testColl.Add(c.Address.RowNumber, c.Address.ColumnNumber, c) |> ignore)
//let testRangeAddress = FsRangeAddress(testHeaderCells.Head.Address, testCells1[testCells1.Length - 1].Address)
//let testTab = FsTable("lel", testRangeAddress)
//testTab |> FsTable.toXlsxTable testColl
//FsRow.
//let testWs = FsWorksheet("testSheet", )

//let toFsWorkbook spreadsheetDoc =
//    let sst = Spreadsheet.tryGetSharedStringTable spreadsheetDoc
//    let wbp = Spreadsheet.getWorkbookPart spreadsheetDoc
//    let wb = Workbook.get wbp
//    let shts = Sheet.Sheets.get wb
//    let shtsN = Sheet.Sheets.getSheets shts |> Array.ofSeq      // N, M, O = multiples of the elements, e.g. here: multiple Sheet elements
//    let wspN =
//        shtsN
//        |> Array.map (
//            fun s -> Worksheet.WorksheetPart.getByID s.Id.Value wbp
//        )
//    let tblNM = wspN |> Array.map (Worksheet.WorksheetPart.getTables >> Array.ofSeq)
//    let wsN = wspN |> Array.map Worksheet.get
//    let sdN = wsN |> Array.map Worksheet.getSheetData
//    let rNM = sdN |> Array.map (SheetData.getRows >> Array.ofSeq)
//    let cNMO =
//        rNM                                             // Level 1: Worksheets (N)
//        |> Array.map (                                  // Level 2: Rows (M)
//            Array.map (Row.toCellSeq >> Array.ofSeq)    // Level 3: Cells (O)
//        )
//    let fswb = new FsWorkbook()
//    let fswsN = 
//        shtsN
//        |> Array.mapi (
//            fun i sht -> 
//                let name = Sheet.getName sht
//                let tblN = tblNM[i]
//                let fsTblN = tblN |> Array.map FsTable.fromXlsxTable
//                let fsCcN
//                FsWorksheet(name, fsRs = , )
//        )
//    0


//let fromXlsxWorksheet xlsxWorksheet =
//    let sd = Worksheet.getSheetData xlsxWorksheet
//    xlsxWorksheet.SheetProperties.LocalName
//    FsWorksheet()
//    0

//doc.Close()