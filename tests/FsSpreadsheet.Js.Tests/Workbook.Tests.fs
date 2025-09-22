module Workbook.Tests

open TestingUtils
open Fable.ExcelJs
open FsSpreadsheet.Js
open FsSpreadsheet

open Fable.Core

module Helper =

    open Fable.Core

    [<Emit("console.log($0)")>]
    let log (obj:obj) = jsNative

open Helper

let private tests_toFsWorkbook = testList "toFsWorkbook" [
    testCase "empty" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.readToFsWorkbook jswb
        Expect.passWithMsg "Convert to fswb"
        let fswsList = fswb.GetWorksheets()
        let jswsList = jswb.worksheets
        Expect.equal fswsList.Count jswsList.Length "Both no worksheet"
    testCase "worksheet" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        let _ = jswb.addWorksheet("My Awesome Worksheet")
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.readToFsWorkbook jswb
        Expect.passWithMsg "Convert to fswb"
        let jswsList = jswb.worksheets
        let fswsList = fswb.GetWorksheets()
        Expect.equal fswsList.Count jswsList.Length "Both 1 worksheet"
        Expect.equal fswsList.[0].Name "My Awesome Worksheet" "Correct worksheet name"
    testCase "worksheets" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        let _ = jswb.addWorksheet("My Awesome Worksheet")
        let _ = jswb.addWorksheet("My Best Worksheet")
        let _ = jswb.addWorksheet("My Nice Worksheet")
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.readToFsWorkbook jswb
        Expect.passWithMsg "Convert to fswb"
        let jswsList = jswb.worksheets
        let fswsList = fswb.GetWorksheets()
        Expect.equal fswsList.Count jswsList.Length "Both 3 worksheets"
        let testWs (index:int) =
            let fsws = fswsList.Item index
            let jsws = jswsList.[index]
            Expect.equal fsws.Name jsws.name $"Correct worksheet name for index {index}."
        for i in 0 .. (jswb.worksheets.Length-1) do
            testWs i
    testCase "table, no body" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        let jsws = jswb.addWorksheet("My Awesome Worksheet")
        let tableColumns = [|TableColumn("Column 1 nice"); TableColumn("Column 2 cool")|]
        let table = Table("My_Awesome_Table", "B1", tableColumns, [||])
        let _ = jsws.addTable(table)
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.readToFsWorkbook jswb
        Expect.passWithMsg "Convert to fswb"
        let fsTables_a = fswb.GetWorksheets().[0].Tables
        let fsTables_b = fswb.GetTables()        
        Expect.hasLength (jsws.getTables()) 1 "js table count"
        Expect.hasLength fsTables_a 1 "fs table count (a)"
        Expect.hasLength fsTables_b 1 "fs table count (b)"
        Expect.equal fsTables_a.[0].Name "My_Awesome_Table" "table name"
        let fsTable = fsTables_a.[0]
        Expect.isTrue (fsTable.ShowHeaderRow) "show header row"
        Expect.equal (fsTable.RangeAddress.Range) ("B1:C1") "RangeAddress"
        let getCellValue (address: string) = 
            let ws = fswb.GetWorksheetAt 1
            let range = FsAddress.fromString(address)
            ws.GetCellAt(range.RowNumber, range.ColumnNumber).Value
            //fsTable.Cell(FsAddress(address), (fswb.GetWorksheetAt 1).CellCollection).Value //bugged in FsSpreadsheet v3.1.1
        Expect.equal (getCellValue "B1") "Column 1 nice" "B1"
        Expect.equal (getCellValue "C1") "Column 2 cool" "C1"

    testCase "table, with body" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        let jsws = jswb.addWorksheet("My Awesome Worksheet")
        let tableColumns = [|TableColumn("Column 1 nice"); TableColumn("Column 2 cool"); TableColumn("Column 3 wow")|]
        let rows = [|
            [| box "Test"; box 2; box true|]
            [| box "Test2"; box 20; box false|]
        |]
        let table = Table("My_Awesome_Table", "B1", tableColumns, rows)
        let _ = jsws.addTable(table)
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.readToFsWorkbook jswb
        Expect.passWithMsg "Convert to fswb"
        let fsTables = fswb.GetTables()        
        Expect.hasLength (jsws.getTables()) 1 "js table count"
        Expect.hasLength fsTables 1 "fs table count "
        let fsTable = fsTables.[0]
        Expect.equal fsTable.Name "My_Awesome_Table" "table name"
        Expect.isTrue (fsTable.ShowHeaderRow) "show header row"
        Expect.equal (fsTable.RangeAddress.Range) ("B1:D3") "RangeAddress"
    testCase "table, with body, check cells" <| fun _ ->
        let jswb = ExcelJs.Excel.Workbook()
        let jsws = jswb.addWorksheet("My Awesome Worksheet")
        let tableColumns = [|TableColumn("Column 1 nice"); TableColumn("Column 2 cool"); TableColumn("Column 3 wow")|]
        let rows = [|
            [| box "Test"; box 2; box true|]
            [| box "Test2"; box 20; box false|]
        |]
        let table = Table("My_Awesome_Table", "B1", tableColumns, rows)
        let _ = jsws.addTable(table)
        Expect.passWithMsg "Create jswb"
        let fswb = JsWorkbook.readToFsWorkbook jswb  
        let getCellValue (address: string) = 
            let ws = fswb.GetWorksheetAt 1
            let range = FsAddress.fromString(address)
            ws.GetCellAt(range.RowNumber, range.ColumnNumber)
            //fsTable.Cell(FsAddress(address), (fswb.GetWorksheetAt 1).CellCollection).Value //bugged in FsSpreadsheet v3.1.1
        let inline expectCellValue (address: string) (getas: FsCell -> 'A) (expectedValue: 'A) =
            let c = getCellValue address
            let a = getas c
            Expect.equal a expectedValue address
        expectCellValue "B1" (fun c -> c.Value) "Column 1 nice"
        expectCellValue "B2" (fun c -> c.Value) "Test"
        expectCellValue "B3" (fun c -> c.Value) "Test2"
        expectCellValue "C1" (fun c -> c.Value) "Column 2 cool"
        expectCellValue "C2" (fun c -> c.ValueAsInt()) 2
        expectCellValue "C3" (fun c -> c.ValueAsInt()) 20
        expectCellValue "D1" (fun c -> c.Value) "Column 3 wow"
        expectCellValue "D2" (fun c -> c.ValueAsBool()) true
        expectCellValue "D3" (fun c -> c.ValueAsBool()) false
]

let tests_toJsWorkbook = testList "toJsWorkbook" [
    testCase "empty" <| fun _ ->
        let fswb = new FsWorkbook()
        Expect.passWithMsg "Create fswb"
        let jswb = JsWorkbook.writeFromFsWorkbook fswb
        Expect.passWithMsg "Convert to jswb"
        let fswsList = fswb.GetWorksheets()
        let jswsList = jswb.worksheets
        Expect.equal fswsList.Count jswsList.Length "Both no worksheet"
    testCase "worksheet" <| fun _ ->
        let fswb = new FsWorkbook()
        let _ = fswb.InitWorksheet("My Awesome Worksheet")
        Expect.passWithMsg "Create fswb"
        let jswb = JsWorkbook.writeFromFsWorkbook fswb
        Expect.passWithMsg "Convert to jswb"
        let fswsList = fswb.GetWorksheets()
        let jswsList = jswb.worksheets
        Expect.equal fswsList.Count jswsList.Length "Both no worksheet"
        Expect.hasLength jswsList 1 "worksheet count"
        Expect.equal jswsList.[0].name "My Awesome Worksheet" "worksheet name"
    testCase "worksheets" <| fun _ ->
        let fswb = new FsWorkbook()
        let _ = fswb.InitWorksheet("My Awesome Worksheet")
        let _ = fswb.InitWorksheet("My cool Worksheet")
        let _ = fswb.InitWorksheet("My wow Worksheet")
        Expect.passWithMsg "Create fswb"
        let jswb = JsWorkbook.writeFromFsWorkbook fswb
        Expect.passWithMsg "Convert to jswb"
        let fswsList = fswb.GetWorksheets()
        let jswsList = jswb.worksheets
        Expect.equal fswsList.Count jswsList.Length "Both no worksheet"
        Expect.hasLength jswsList 3 "worksheet count"
        Expect.equal jswsList.[0].name "My Awesome Worksheet" "1 worksheet name"
        Expect.equal jswsList.[1].name "My cool Worksheet" "2 worksheet name"
        Expect.equal jswsList.[2].name "My wow Worksheet" "3 worksheet name"
    testCase "table, no body" <| fun _ ->
        let fswb = new FsWorkbook()
        let fsws = fswb.InitWorksheet("My Awesome Worksheet")
        let _ = fsws.AddCell(FsCell("My Column 1",address=FsAddress.fromString("B1")))
        let _ = fsws.AddCell(FsCell("My Column 2",address=FsAddress.fromString("C1")))
        let t = FsTable("My_New_Table", FsRangeAddress.fromString("B1:C1"))
        let _ = fsws.AddTable(t)
        Expect.passWithMsg "Create jswb"
        let jswb = JsWorkbook.writeFromFsWorkbook fswb
        Expect.passWithMsg "Convert to fswb"
        let jsws = jswb.worksheets.[0]
        Expect.equal jsws.name "My Awesome Worksheet" "ws name"
        let jstables = jsws.getTables()
        Expect.hasLength jstables 1 "table count"
        let jstableref = jstables.[0]
        Expect.equal jstableref.name "My_New_Table" "table name"
        let jstable = jstables.[0].table.Value
        Expect.hasLength jstable.rows 0 "table body is 0"
        Expect.equal jstable.columns.[0].name "My Column 1" "My Column 1"
        Expect.equal jstable.columns.[1].name "My Column 2" "My Column 2"
    testCase "table, with body" <| fun _ ->
        let fswb = new FsWorkbook()
        let fsws = fswb.InitWorksheet("My Awesome Worksheet")
        let _ = fsws.AddCell(FsCell("My Column 1",address=FsAddress.fromString("B1")))
        let _ = fsws.AddCell(FsCell(2,address=FsAddress.fromString("B2")))
        let _ = fsws.AddCell(FsCell(20,address=FsAddress.fromString("B3")))
        let _ = fsws.AddCell(FsCell("My Column 2",address=FsAddress.fromString("C1")))
        let _ = fsws.AddCell(FsCell("row2",address=FsAddress.fromString("C2")))
        let _ = fsws.AddCell(FsCell("row20",address=FsAddress.fromString("C3")))
        let _ = fsws.AddCell(FsCell("My Column 3",address=FsAddress.fromString("D1")))
        let _ = fsws.AddCell(FsCell(true,address=FsAddress.fromString("D2")))
        let _ = fsws.AddCell(FsCell(false,address=FsAddress.fromString("D3")))
        let t = FsTable("My_New_Table", FsRangeAddress.fromString("B1:D3"))
        let _ = fsws.AddTable(t)
        Expect.passWithMsg "Create jswb"
        let jswb = JsWorkbook.writeFromFsWorkbook fswb
        Expect.passWithMsg "Convert to fswb"
        let jsws = jswb.worksheets.[0]
        Expect.equal jsws.name "My Awesome Worksheet" "ws name"
        let jstables = jsws.getTables()
        Expect.hasLength jstables 1 "table count"
        let jstableref = jstables.[0]
        Expect.equal jstableref.name "My_New_Table" "table name"
        let jstable = jstables.[0].table.Value
        Expect.equal jstable.columns.[0].name "My Column 1" "My Column 1"
        Expect.equal jstable.columns.[1].name "My Column 2" "My Column 2"
        Expect.equal jstable.columns.[2].name "My Column 3" "My Column 3"
        Expect.hasLength jstable.rows 2 "table body has 2 rows"
    testCase "table, with body, check cells" <| fun _ ->
        let fswb = new FsWorkbook()
        let fsws = fswb.InitWorksheet("My Awesome Worksheet")
        let _ = fsws.AddCell(FsCell("My Column 1",address=FsAddress.fromString("B1")))
        let _ = fsws.AddCell(FsCell(2,DataType.Number,address=FsAddress.fromString("B2")))
        let _ = fsws.AddCell(FsCell(20,DataType.Number,address=FsAddress.fromString("B3")))
        let _ = fsws.AddCell(FsCell("My Column 2",address=FsAddress.fromString("C1")))
        let _ = fsws.AddCell(FsCell("row2",address=FsAddress.fromString("C2")))
        let _ = fsws.AddCell(FsCell("row20",address=FsAddress.fromString("C3")))
        let _ = fsws.AddCell(FsCell("My Column 3",address=FsAddress.fromString("D1")))
        let _ = fsws.AddCell(FsCell(true, DataType.Boolean, address=FsAddress.fromString("D2")))
        let _ = fsws.AddCell(FsCell(false, DataType.Boolean, address=FsAddress.fromString("D3")))
        let t = FsTable("My_New_Table", FsRangeAddress.fromString("B1:D3"))
        let _ = fsws.AddTable(t)
        fsws.RescanRows()
        Expect.passWithMsg "Create jswb"
        let jswb = JsWorkbook.writeFromFsWorkbook fswb
        let jstable = jswb.worksheets.[0].getTables().[0].table.Value
        let row0 = jstable.rows.[0]
        let row1 = jstable.rows.[1]
        Expect.hasLength row0 3 "row B2:D2 length"
        Expect.hasLength row1 3 "row B3:D3 length"
        Expect.equal row0.[0] 2 "B2"
        Expect.equal row1.[0] 20 "B3"
        Expect.equal row0.[1] "row2" "C2"
        Expect.equal row1.[1] "row20" "C3"
        Expect.equal row0.[2] true "D2"
        Expect.equal row1.[2] false "D3"
    testCase "table ColumnCount()" <| fun _ ->
        let fswb = new FsWorkbook()
        let fsws = fswb.InitWorksheet("My Awesome Worksheet")
        let _ = fsws.AddCell(FsCell("My Column 1",address=FsAddress.fromString("B1")))
        let _ = fsws.AddCell(FsCell(2,DataType.Number,address=FsAddress.fromString("B2")))
        let _ = fsws.AddCell(FsCell(20,DataType.Number,address=FsAddress.fromString("B3")))
        let _ = fsws.AddCell(FsCell("My Column 2",address=FsAddress.fromString("C1")))
        let _ = fsws.AddCell(FsCell("row2",address=FsAddress.fromString("C2")))
        let _ = fsws.AddCell(FsCell("row20",address=FsAddress.fromString("C3")))
        let _ = fsws.AddCell(FsCell("My Column 3",address=FsAddress.fromString("D1")))
        let _ = fsws.AddCell(FsCell(true, DataType.Boolean, address=FsAddress.fromString("D2")))
        let _ = fsws.AddCell(FsCell(false, DataType.Boolean, address=FsAddress.fromString("D3")))
        let t = FsTable("My_New_Table", FsRangeAddress.fromString("B1:D3"))
        let _ = fsws.AddTable(t)
        fsws.RescanRows()
        let table = fsws.Tables.[0]
        Expect.equal (table.ColumnCount()) 3 "column count"
    testCase "table with implicit set" <| fun _ ->
        let fswb = new FsWorkbook()
        let fsws = fswb.InitWorksheet("My Awesome Worksheet")
        fsws.Row(1).[2].SetValueAs "My Column 1"
        fsws.Row(2).[2].SetValueAs 2
        fsws.Row(3).[2].SetValueAs 20
        fsws.Row(1).[3].SetValueAs "My Column 2"
        fsws.Row(2).[3].SetValueAs "row2"
        fsws.Row(3).[3].SetValueAs "row20"
        fsws.Row(1).[4].SetValueAs "My Column 2"
        fsws.Row(2).[4].SetValueAs true
        fsws.Row(3).[4].SetValueAs false
        let t = FsTable("My New Table", FsRangeAddress.fromString("B1:D3"))
        let _ = fsws.AddTable(t)
        fsws.RescanRows()
        let table = fsws.Tables.[0]
        Expect.equal (table.ColumnCount()) 3 "column count"
]

open Fable.Core

let tests_xlsx = testList "xlsx" [
    testList "read" [
        testCaseAsync "isa.assay.xlsx" <| async {
            let! fswb = Xlsx.fromXlsxFile("./tests/JS/TestFiles/isa.assay.xlsx") |> Async.AwaitPromise
            Expect.equal (fswb.GetWorksheets().Count) 5 "Count"
        }
    ]
]

let performance =
    testList "Performace" [
        testCaseAsync "ReadBigFile" <| async {
            let sw = Stopwatch()        
            let p = DefaultTestObject.BigFile.asRelativePathNode
            sw.Start()
            let! wb = FsWorkbook.fromXlsxFile(p) |> Async.AwaitPromise
            sw.Stop()
            let ms = sw.Elapsed.Milliseconds
            Expect.isTrue (ms < 2000)  $"Elapsed time should be less than 2000ms but was {ms}ms"
            Expect.equal (wb.GetWorksheetAt(1).Rows.Count) 153991 "Row count should be 153991"
        }
    ]

let main = testList "JsWorkbook<->FsWorkbook" [
    tests_toFsWorkbook
    tests_toJsWorkbook
    tests_xlsx
    performance
]

