module DefaultTestObject

open FsSpreadsheet
open Fable.Core
#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif

let [<Literal>] testFolder = "TestFiles"

[<AttachMembers>]
type TestFiles =
| Excel
| Libre
| FableExceljs
| ClosedXML
| FsSpreadsheet

    member this.asFileName =
        match this with
        | Excel             -> "TestWorkbook_Excel.xlsx"
        | Libre             -> "TestWorkbook_Libre.xlsx"
        | FableExceljs      -> "TestWorkbook_FableExcelJS.xlsx"
        | ClosedXML         -> "TestWorkbook_ClosedXML.xlsx"
        | FsSpreadsheet     -> "TestWorkbook_FsSpreadsheet.xlsx"

    member this.asRelativePath = $"../TestUtils/{testFolder}/{this.asFileName}"

module ExpectedRows = 
    let headerRow (range:string) cc = 
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs "Numbers"
        row[2].SetValueAs "Strings"
        row[3].SetValueAs "DateTime"
        row[4].SetValueAs "ARCtrl Column"
        row[5].SetValueAs "ARCtrl Column "
        row
    let firstRow(range: string) cc = 
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 1
        row[2].SetValueAs "Hello"
        row[3].SetValueAs (System.DateTime(2023,10,14))
        row[4].SetValueAs "(A) This is part 1 of 2"
        row[5].SetValueAs "(A) This is part 2 of 2"
        row
    let secondRow(range:string) cc =
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 2
        row[2].SetValueAs "World"
        row[3].SetValueAs (System.DateTime(2023,10,15))
        row[5].SetValueAs "Tests if column names with whitespace at end can be unique"
        row
    let thirdRow(range:string) cc =
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 3
        row[2].SetValueAs "Bye"
        row[3].SetValueAs (System.DateTime(2023,10,16))
        row
    let fourthRow(range:string) cc = 
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 4
        row[2].SetValueAs "Outer Space"
        row[3].SetValueAs (System.DateTime(2023,10,17))
        row

    let rowCollectionA1 = 
        let cells = FsCellsCollection() 
        [|headerRow("A1:E1") ;firstRow("A2:E2") ;secondRow("A3:E3");thirdRow("A4:E4");fourthRow("A5:E5")|] 
        |> Array.map (fun x -> x cells)
    let rowCollectionB4 = 
        let cells = FsCellsCollection() 
        [|headerRow("B4:E4");firstRow("B5:E5");secondRow("B6:E6");thirdRow("B7:E7");fourthRow("B8:E8")|]
        |> Array.map (fun x -> x cells)
            
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
        Sheet1.sheetName, (Some Sheet1.tableName, ExpectedRows.rowCollectionA1); 
        Sheet2.sheetName, (None, ExpectedRows.rowCollectionA1);       
        Sheet3.sheetName, (Some Sheet3.tableName, ExpectedRows.rowCollectionB4)                  
    ]
    |> Map.ofList
