module Workbook.Tests

open TestingUtils
open Fable.Openpyxl
open FsSpreadsheet.Py
open FsSpreadsheet

open Fable.Core.PyInterop
open Fable.Core
open Fable.Pyxpecto

let private tests_toFsWorkbook = testList "toFsWorkbook" [
    testCase "empty" <| fun _ ->
        let pyWB = Workbook.create()
        let fsWB = PyWorkbook.toFsWorkbook pyWB
        
        Expect.equal (fsWB.GetWorksheets().Count) 0 "both no worksheet"
    //testCase "worksheet" <| fun _ ->
    //    let jswb = exceljs.excel.workbook()
    //    let _ = jswb.addworksheet("my awesome worksheet")
    //    expect.passwithmsg "create jswb"
    //    let fswb = jsworkbook.readtofsworkbook jswb
    //    expect.passwithmsg "convert to fswb"
    //    let jswslist = jswb.worksheets
    //    let fswslist = fswb.getworksheets()
    //    expect.equal fswslist.count jswslist.length "both 1 worksheet"
    //    expect.equal fswslist.[0].name "my awesome worksheet" "correct worksheet name"
    //testCase "worksheets" <| fun _ ->
    //    let jswb = exceljs.excel.workbook()
    //    let _ = jswb.addworksheet("my awesome worksheet")
    //    let _ = jswb.addworksheet("my best worksheet")
    //    let _ = jswb.addworksheet("my nice worksheet")
    //    expect.passwithmsg "create jswb"
    //    let fswb = jsworkbook.readtofsworkbook jswb
    //    expect.passwithmsg "convert to fswb"
    //    let jswslist = jswb.worksheets
    //    let fswslist = fswb.getworksheets()
    //    expect.equal fswslist.count jswslist.length "both 3 worksheets"
    //    let testws (index:int) =
    //        let fsws = fswslist.item index
    //        let jsws = jswslist.[index]
    //        expect.equal fsws.name jsws.name $"correct worksheet name for index {index}."
    //    for i in 0 .. (jswb.worksheets.length-1) do
    //        testws i
    //testCase "table, no body" <| fun _ ->
    //    let jswb = exceljs.excel.workbook()
    //    let jsws = jswb.addworksheet("my awesome worksheet")
    //    let tablecolumns = [|tablecolumn("column 1 nice"); tablecolumn("column 2 cool")|]
    //    let table = table("my_awesome_table", "b1", tablecolumns, [||])
    //    let _ = jsws.addtable(table)
    //    expect.passwithmsg "create jswb"
    //    let fswb = jsworkbook.readtofsworkbook jswb
    //    expect.passwithmsg "convert to fswb"
    //    let fstables_a = fswb.getworksheets().[0].tables
    //    let fstables_b = fswb.gettables()        
    //    expect.haslength (jsws.gettables()) 1 "js table count"
    //    expect.haslength fstables_a 1 "fs table count (a)"
    //    expect.haslength fstables_b 1 "fs table count (b)"
    //    expect.equal fstables_a.[0].name "my_awesome_table" "table name"
    //    let fstable = fstables_a.[0]
    //    expect.istrue (fstable.showheaderrow) "show header row"
    //    expect.equal (fstable.rangeaddress.range) ("b1:c1") "rangeaddress"
    //    let getcellvalue (address: string) = 
    //        let ws = fswb.getworksheetat 1
    //        let range = fsaddress(address)
    //        ws.getcellat(range.rownumber, range.columnnumber).value
    //        //fstable.cell(fsaddress(address), (fswb.getworksheetat 1).cellcollection).value //bugged in fsspreadsheet v3.1.1
    //    expect.equal (getcellvalue "b1") "column 1 nice" "b1"
    //    expect.equal (getcellvalue "c1") "column 2 cool" "c1"

    //testCase "table, with body" <| fun _ ->
    //    let jswb = exceljs.excel.workbook()
    //    let jsws = jswb.addworksheet("my awesome worksheet")
    //    let tablecolumns = [|tablecolumn("column 1 nice"); tablecolumn("column 2 cool"); tablecolumn("column 3 wow")|]
    //    let rows = [|
    //        [| box "test"; box 2; box true|]
    //        [| box "test2"; box 20; box false|]
    //    |]
    //    let table = table("my_awesome_table", "b1", tablecolumns, rows)
    //    let _ = jsws.addtable(table)
    //    expect.passwithmsg "create jswb"
    //    let fswb = jsworkbook.readtofsworkbook jswb
    //    expect.passwithmsg "convert to fswb"
    //    let fstables = fswb.gettables()        
    //    expect.haslength (jsws.gettables()) 1 "js table count"
    //    expect.haslength fstables 1 "fs table count "
    //    let fstable = fstables.[0]
    //    expect.equal fstable.name "my_awesome_table" "table name"
    //    expect.istrue (fstable.showheaderrow) "show header row"
    //    expect.equal (fstable.rangeaddress.range) ("b1:d3") "rangeaddress"
    //testCase "table, with body, check cells" <| fun _ ->
    //    let jswb = exceljs.excel.workbook()
    //    let jsws = jswb.addworksheet("my awesome worksheet")
    //    let tablecolumns = [|tablecolumn("column 1 nice"); tablecolumn("column 2 cool"); tablecolumn("column 3 wow")|]
    //    let rows = [|
    //        [| box "test"; box 2; box true|]
    //        [| box "test2"; box 20; box false|]
    //    |]
    //    let table = table("my_awesome_table", "b1", tablecolumns, rows)
    //    let _ = jsws.addtable(table)
    //    expect.passwithmsg "create jswb"
    //    let fswb = jsworkbook.readtofsworkbook jswb  
    //    let getcellvalue (address: string) = 
    //        let ws = fswb.getworksheetat 1
    //        let range = fsaddress(address)
    //        ws.getcellat(range.rownumber, range.columnnumber)
    //        //fstable.cell(fsaddress(address), (fswb.getworksheetat 1).cellcollection).value //bugged in fsspreadsheet v3.1.1
    //    let inline expectcellvalue (address: string) (getas: fscell -> 'a) (expectedvalue: 'a) =
    //        let c = getcellvalue address
    //        let a = getas c
    //        expect.equal a expectedvalue address
    //    expectcellvalue "b1" (fun c -> c.value) "column 1 nice"
    //    expectcellvalue "b2" (fun c -> c.value) "test"
    //    expectcellvalue "b3" (fun c -> c.value) "test2"
    //    expectcellvalue "c1" (fun c -> c.value) "column 2 cool"
    //    expectcellvalue "c2" (fun c -> c.valueasint()) 2
    //    expectcellvalue "c3" (fun c -> c.valueasint()) 20
    //    expectcellvalue "d1" (fun c -> c.value) "column 3 wow"
    //    expectcellvalue "d2" (fun c -> c.valueasbool()) true
    //    expectcellvalue "d3" (fun c -> c.valueasbool()) false
]

let tests_toPyWorkbook = testList "toPyWorkbook" [
    testCase "empty" <| fun _ ->
        
        let fsWB = new FsWorkbook()
        Expect.fails (fun () -> PyWorkbook.fromFsWorkbook fsWB |> ignore) "no worksheet given"
    testCase "worksheet" <| fun _ ->
        let fsWB = new FsWorkbook()
        let _ = fsWB.InitWorksheet("my awesome worksheet")
        let pyWB = PyWorkbook.fromFsWorkbook fsWB
        Expect.passWithMsg "convert to jswb"
        let pyWsList = pyWB?worksheets
        Expect.equal (pyWsList |> Array.length) 1 "worksheet count"
    testCase "worksOnFilledTable (issue #100)" <| fun _ ->
        let fsWB = new FsWorkbook()
        let ws = fsWB.InitWorksheet("my awesome worksheet")
        ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("b1")) "my column 1") |> ignore
        ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("c1")) "my column 2") |> ignore
        ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("b2")) 2) |> ignore
        ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("c2")) "row2") |> ignore
        ws.AddTable(FsTable("my_new_table", FsRangeAddress.fromString("b1:c2"))) |> ignore
        PyWorkbook.fromFsWorkbook fsWB |> ignore
    testCase "failsOnEmptyTable (issue #100)" <| fun _ ->
        let fsWB = new FsWorkbook()
        let ws = fsWB.InitWorksheet("my awesome worksheet")
        ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("b1")) "my column 1") |> ignore
        ws.AddCell(FsCell.createWithAdress(FsAddress.fromString("c1")) "my column 2") |> ignore
        ws.AddTable(FsTable("my_new_table", FsRangeAddress.fromString("b1:c1"))) |> ignore
        Expect.throws (fun () -> PyWorkbook.fromFsWorkbook fsWB |> ignore) "no body in table"
        
        //pyWB?worksheets
        //|> List.map (fun (ws : FsWorksheet) -> ws.Name = "my awesome worksheet")


    //testCase "worksheets" <| fun _ ->
    //    let fswb = new fsworkbook()
    //    let _ = fswb.initworksheet("my awesome worksheet")
    //    let _ = fswb.initworksheet("my cool worksheet")
    //    let _ = fswb.initworksheet("my wow worksheet")
    //    expect.passwithmsg "create fswb"
    //    let jswb = jsworkbook.writefromfsworkbook fswb
    //    expect.passwithmsg "convert to jswb"
    //    let fswslist = fswb.getworksheets()
    //    let jswslist = jswb.worksheets
    //    expect.equal fswslist.count jswslist.length "both no worksheet"
    //    expect.haslength jswslist 3 "worksheet count"
    //    expect.equal jswslist.[0].name "my awesome worksheet" "1 worksheet name"
    //    expect.equal jswslist.[1].name "my cool worksheet" "2 worksheet name"
    //    expect.equal jswslist.[2].name "my wow worksheet" "3 worksheet name"
    //testCase "table, no body" <| fun _ ->
    //    let fswb = new fsworkbook()
    //    let fsws = fswb.initworksheet("my awesome worksheet")
    //    let _ = fsws.addcell(fscell("my column 1",address=fsaddress("b1")))
    //    let _ = fsws.addcell(fscell("my column 2",address=fsaddress("c1")))
    //    let t = fstable("my_new_table", fsrangeaddress("b1:c1"))
    //    let _ = fsws.addtable(t)
    //    expect.passwithmsg "create jswb"
    //    let jswb = jsworkbook.writefromfsworkbook fswb
    //    expect.passwithmsg "convert to fswb"
    //    let jsws = jswb.worksheets.[0]
    //    expect.equal jsws.name "my awesome worksheet" "ws name"
    //    let jstables = jsws.gettables()
    //    expect.haslength jstables 1 "table count"
    //    let jstableref = jstables.[0]
    //    expect.equal jstableref.name "my_new_table" "table name"
    //    let jstable = jstables.[0].table.value
    //    expect.haslength jstable.rows 0 "table body is 0"
    //    expect.equal jstable.columns.[0].name "my column 1" "my column 1"
    //    expect.equal jstable.columns.[1].name "my column 2" "my column 2"
    //testCase "table, with body" <| fun _ ->
    //    let fswb = new fsworkbook()
    //    let fsws = fswb.initworksheet("my awesome worksheet")
    //    let _ = fsws.addcell(fscell("my column 1",address=fsaddress("b1")))
    //    let _ = fsws.addcell(fscell(2,address=fsaddress("b2")))
    //    let _ = fsws.addcell(fscell(20,address=fsaddress("b3")))
    //    let _ = fsws.addcell(fscell("my column 2",address=fsaddress("c1")))
    //    let _ = fsws.addcell(fscell("row2",address=fsaddress("c2")))
    //    let _ = fsws.addcell(fscell("row20",address=fsaddress("c3")))
    //    let _ = fsws.addcell(fscell("my column 3",address=fsaddress("d1")))
    //    let _ = fsws.addcell(fscell(true,address=fsaddress("d2")))
    //    let _ = fsws.addcell(fscell(false,address=fsaddress("d3")))
    //    let t = fstable("my_new_table", fsrangeaddress("b1:d3"))
    //    let _ = fsws.addtable(t)
    //    expect.passwithmsg "create jswb"
    //    let jswb = jsworkbook.writefromfsworkbook fswb
    //    expect.passwithmsg "convert to fswb"
    //    let jsws = jswb.worksheets.[0]
    //    expect.equal jsws.name "my awesome worksheet" "ws name"
    //    let jstables = jsws.gettables()
    //    expect.haslength jstables 1 "table count"
    //    let jstableref = jstables.[0]
    //    expect.equal jstableref.name "my_new_table" "table name"
    //    let jstable = jstables.[0].table.value
    //    expect.equal jstable.columns.[0].name "my column 1" "my column 1"
    //    expect.equal jstable.columns.[1].name "my column 2" "my column 2"
    //    expect.equal jstable.columns.[2].name "my column 3" "my column 3"
    //    expect.haslength jstable.rows 2 "table body has 2 rows"
    //testCase "table, with body, check cells" <| fun _ ->
    //    let fswb = new fsworkbook()
    //    let fsws = fswb.initworksheet("my awesome worksheet")
    //    let _ = fsws.addcell(fscell("my column 1",address=fsaddress("b1")))
    //    let _ = fsws.addcell(fscell(2,datatype.number,address=fsaddress("b2")))
    //    let _ = fsws.addcell(fscell(20,datatype.number,address=fsaddress("b3")))
    //    let _ = fsws.addcell(fscell("my column 2",address=fsaddress("c1")))
    //    let _ = fsws.addcell(fscell("row2",address=fsaddress("c2")))
    //    let _ = fsws.addcell(fscell("row20",address=fsaddress("c3")))
    //    let _ = fsws.addcell(fscell("my column 3",address=fsaddress("d1")))
    //    let _ = fsws.addcell(fscell(true, datatype.boolean, address=fsaddress("d2")))
    //    let _ = fsws.addcell(fscell(false, datatype.boolean, address=fsaddress("d3")))
    //    let t = fstable("my_new_table", fsrangeaddress("b1:d3"))
    //    let _ = fsws.addtable(t)
    //    fsws.rescanrows()
    //    expect.passwithmsg "create jswb"
    //    let jswb = jsworkbook.writefromfsworkbook fswb
    //    let jstable = jswb.worksheets.[0].gettables().[0].table.value
    //    let row0 = jstable.rows.[0]
    //    let row1 = jstable.rows.[1]
    //    expect.haslength row0 3 "row b2:d2 length"
    //    expect.haslength row1 3 "row b3:d3 length"
    //    expect.equal row0.[0] 2 "b2"
    //    expect.equal row1.[0] 20 "b3"
    //    expect.equal row0.[1] "row2" "c2"
    //    expect.equal row1.[1] "row20" "c3"
    //    expect.equal row0.[2] true "d2"
    //    expect.equal row1.[2] false "d3"
    //testCase "table columncount()" <| fun _ ->
    //    let fswb = new fsworkbook()
    //    let fsws = fswb.initworksheet("my awesome worksheet")
    //    let _ = fsws.addcell(fscell("my column 1",address=fsaddress("b1")))
    //    let _ = fsws.addcell(fscell(2,datatype.number,address=fsaddress("b2")))
    //    let _ = fsws.addcell(fscell(20,datatype.number,address=fsaddress("b3")))
    //    let _ = fsws.addcell(fscell("my column 2",address=fsaddress("c1")))
    //    let _ = fsws.addcell(fscell("row2",address=fsaddress("c2")))
    //    let _ = fsws.addcell(fscell("row20",address=fsaddress("c3")))
    //    let _ = fsws.addcell(fscell("my column 3",address=fsaddress("d1")))
    //    let _ = fsws.addcell(fscell(true, datatype.boolean, address=fsaddress("d2")))
    //    let _ = fsws.addcell(fscell(false, datatype.boolean, address=fsaddress("d3")))
    //    let t = fstable("my_new_table", fsrangeaddress("b1:d3"))
    //    let _ = fsws.addtable(t)
    //    fsws.rescanrows()
    //    let table = fsws.tables.[0]
    //    expect.equal (table.columncount()) 3 "column count"
    //testCase "table with implicit set" <| fun _ ->
    //    let fswb = new fsworkbook()
    //    let fsws = fswb.initworksheet("my awesome worksheet")
    //    fsws.row(1).[2].setvalueas "my column 1"
    //    fsws.row(2).[2].setvalueas 2
    //    fsws.row(3).[2].setvalueas 20
    //    fsws.row(1).[3].setvalueas "my column 2"
    //    fsws.row(2).[3].setvalueas "row2"
    //    fsws.row(3).[3].setvalueas "row20"
    //    fsws.row(1).[4].setvalueas "my column 2"
    //    fsws.row(2).[4].setvalueas true
    //    fsws.row(3).[4].setvalueas false
    //    let t = fstable("my new table", fsrangeaddress("b1:d3"))
    //    let _ = fsws.addtable(t)
    //    fsws.rescanrows()
    //    let table = fsws.tables.[0]
    //    expect.equal (table.columncount()) 3 "column count"
]

//open fable.core

//let tests_xlsx = testList "xlsx" [
//    testList "read" [
//        testCaseasync "isa.assay.xlsx" <| async {
//            let! fswb = xlsx.fromxlsxfile("./tests/js/testfiles/isa.assay.xlsx") |> async.awaitpromise
//            expect.equal (fswb.getworksheets().count) 5 "count"
//        }
//    ]
//]

//let performance =
//    testList "performace" [
//        testCaseasync "readbigfile" <| async {
//            let sw = stopwatch()        
//            let p = defaulttestobject.bigfile.asrelativepathnode
//            sw.start()
//            let! wb = fsworkbook.fromxlsxfile(p) |> async.awaitpromise
//            sw.stop()
//            let ms = sw.elapsed.milliseconds
//            expect.istrue (ms < 2000)  $"elapsed time should be less than 2000ms but was {ms}ms"
//            expect.equal (wb.getworksheetat(1).rows.count) 153991 "row count should be 153991"
//        }
//    ]


let main = testList "Workbook" [
    tests_toFsWorkbook
    tests_toPyWorkbook
    //tests_xlsx
    //performance
]

