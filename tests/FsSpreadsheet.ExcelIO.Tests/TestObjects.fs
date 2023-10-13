module TestObjects

open FsSpreadsheet

module CrossIOTests = 
    
    let testFolder = "TestFiles"

    let excelFileName = "TestWorkbook_Excel.xlsx"
    let libreFileName = "TestWorkbook_Libre.xlsx"
    let exceljsFileName = "TestWorkbook_ExcelJS.xlsx"
    let closesXMLFileName = "TestWorkbook_ClosedXML.xlsx"
    let FsSpreadsheetFileName = "TestWorkbook_FsSpreadsheet.xlsx"

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


let sheet1Name = "MySheet1"
let sheet2Name = "MySheet2"

let sheet1() =
    let ws = new FsWorksheet(sheet1Name)
    [
        FsCell.createWithDataType DataType.String 1 1 "A1"
        FsCell.createWithDataType DataType.String 1 2 "B1"
        FsCell.createWithDataType DataType.String 1 3 "C1"

        FsCell.createWithDataType DataType.String 2 1 "A2"
        FsCell.createWithDataType DataType.String 2 2 "B2"
        FsCell.createWithDataType DataType.String 2 3 "C2"

        FsCell.createWithDataType DataType.String 3 1 "A3"
        FsCell.createWithDataType DataType.String 3 2 "B3"
        FsCell.createWithDataType DataType.String 3 3 "C3"        
    ]
    |> List.iter (fun c -> ws.Row(c.RowNumber).[c.ColumnNumber].SetValueAs c.Value)
    ws

let sheet2() =
    let ws = new FsWorksheet(sheet2Name)
    [
        FsCell.createWithDataType DataType.Number 1 1 1
        FsCell.createWithDataType DataType.Number 1 2 2
        FsCell.createWithDataType DataType.Number 1 3 3
        FsCell.createWithDataType DataType.Number 1 4 4

        FsCell.createWithDataType DataType.Number 2 1 5
        FsCell.createWithDataType DataType.Number 2 2 6
        FsCell.createWithDataType DataType.Number 2 3 7
        FsCell.createWithDataType DataType.Number 2 4 8     
    ]
    |> List.iter (fun c -> ws.Row(c.RowNumber).[c.ColumnNumber].SetValueAs c.Value)
    ws