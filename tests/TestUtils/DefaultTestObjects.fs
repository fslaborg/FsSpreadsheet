module TestUtils.DefaultTestObject

open FsSpreadsheet
#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif

let [<Literal>] testFolder = "TestFiles"
let [<Literal>] excelFileName = "TestWorkbook_Excel.xlsx"
let [<Literal>] libreFileName = "TestWorkbook_Libre.xlsx"
let [<Literal>] exceljsFileName = "TestWorkbook_ExcelJS.xlsx"
let [<Literal>] closesXMLFileName = "TestWorkbook_ClosedXML.xlsx"
let [<Literal>] FsSpreadsheetFileName = "TestWorkbook_FsSpreadsheet.xlsx"

module ExpectedRows = 
    let cells = FsCellsCollection()
    let headerRow = 
        let row = FsRow(FsRangeAddress("A1:E1"),cells)
        row[1].SetValueAs "Numbers"
        row[2].SetValueAs "Strings"
        row[3].SetValueAs "DateTime"
        row[4].SetValueAs "ARCtrl Column"
        row[5].SetValueAs "ARCtrl Column "
        row
    let firstRow = 
        let row = FsRow(FsRangeAddress("A2:E2"),cells)
        row[1].SetValueAs 1
        row[2].SetValueAs "Hello"
        row[3].SetValueAs (System.DateTime(2023,10,14))
        row[4].SetValueAs "(A) This is part 1 of 2"
        row[5].SetValueAs "(A) This is part 2 of 2"
        row
    let secondRow =
        let row = FsRow(FsRangeAddress("A3:E3"),cells)
        row[1].SetValueAs 2
        row[2].SetValueAs "World"
        row[3].SetValueAs (System.DateTime(2023,10,15))
        row[5].SetValueAs "Tests if column names with whitespace at end can be unique"
        row
    let thirdRow =
        let row = FsRow(FsRangeAddress("A4:E4"),cells)
        row[1].SetValueAs 3
        row[2].SetValueAs "Bye"
        row[3].SetValueAs (System.DateTime(2023,10,16))
        row
    let fourthRow = 
        let row = FsRow(FsRangeAddress("A5:E5"),cells)
        row[1].SetValueAs 4
        row[2].SetValueAs "Outer Space"
        row[3].SetValueAs (System.DateTime(2023,10,17))
        row

    let rowCollection =
        {|HeaderRow = headerRow; Body = [|firstRow;secondRow;thirdRow;fourthRow|]|}
            

module Sheet1 = 

    [<Literal>]
    let sheetName = "WithTable"
    [<Literal>]
    let tableName = "MyTable"

module Sheet2 = 

    [<Literal>]
    let sheetName = "Tableless"

module Sheet3 = 

    [<Literal>]
    let sheetName = "WithTable_Duplicate"
    [<Literal>]
    let tableName = "MyOtherTable"

let valueMap = 
    [
        Sheet1.sheetName, (Some Sheet1.tableName, ExpectedRows.rowCollection); 
        Sheet2.sheetName, (None, ExpectedRows.rowCollection);       
        Sheet3.sheetName, (Some Sheet3.tableName, ExpectedRows.rowCollection)                  
    ]
    |> Map.ofList

let isDefaultTestObject (wb: FsWorkbook) =
    let worksheets = wb.GetWorksheets()
    for ws in worksheets do
        let isTable, wsInfo = Expect.wantSome (valueMap |> Map.tryFind ws.Name) $"ExpectError: Unable to get info for worksheet: {ws.Name}"
        match isTable with
        | Some expectedTableName -> 
            let actualTable = Expect.wantSome (ws.Tables |> Seq.tryFind (fun t -> t.Name = expectedTableName)) $"ExpectError: Unable to get info for worksheet->table: {ws.Name}->{expectedTableName}"
            let headerRow = Expect.wantSome (actualTable.TryGetHeaderRow(ws.CellCollection)) $"ExpectError: ShowHeaderRow is false for worksheet->table: {ws.Name}->{expectedTableName}"
            let actualRows = actualTable.GetRows(ws.CellCollection) |> Seq.tail //Seq.tail skips HeaderRow
            Expect.equal headerRow wsInfo.HeaderRow $"ExpectError: HeaderRow is not equal for worksheet->table: {ws.Name}->{expectedTableName}"
            Expect.sequenceEqual actualRows wsInfo.Body $"ExpectError: Table body rows are not equal for worksheet->table: {ws.Name}->{expectedTableName}"
        | None ->
            let actualRows = ws.Rows
            Expect.sequenceEqual actualRows wsInfo.Body $"ExpectError: Table body rows are not equal for worksheet: {ws.Name}"

