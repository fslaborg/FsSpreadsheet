﻿module DefaultTestObject

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
| FsSpreadsheetNET
| FsSpreadsheetJS

    member this.asFileName =
        match this with
        | Excel             -> "TestWorkbook_Excel.xlsx"
        | Libre             -> "TestWorkbook_Libre.xlsx"
        | FableExceljs      -> "TestWorkbook_FableExceljs.xlsx"
        | ClosedXML         -> "TestWorkbook_ClosedXML.xlsx"
        | FsSpreadsheetNET  -> "TestWorkbook_FsSpreadsheet.net.xlsx"
        | FsSpreadsheetJS  -> "TestWorkbook_FsSpreadsheet.js.xlsx"

    member this.asRelativePath = $"../TestUtils/{testFolder}/{this.asFileName}"
    member this.asRelativePathNode = $"./tests/TestUtils/{testFolder}/{this.asFileName}"

module ExpectedRows = 
    let headerRow (range:string) cc = 
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs "Numbers"
        row[2].SetValueAs "Strings"
        row[3].SetValueAs "DateTime"
        row[4].SetValueAs "Boolean"
        row[5].SetValueAs "ARCtrl Column"
        row[6].SetValueAs "ARCtrl Column "
        row
    let firstRow(range: string) cc = 
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 1.
        row[2].SetValueAs "Hello"
        row[3].SetValueAs (System.DateTime(2023,10,14,0,0,0))
        row[4].SetValueAs<bool> true
        row[5].SetValueAs "(A) This is part 1 of 2"
        row[6].SetValueAs "(A) This is part 2 of 2"
        row
    let secondRow(range:string) cc =
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 2.
        row[2].SetValueAs "World"
        row[3].SetValueAs (System.DateTime(2023,10,15, 18,0,0))
        row[4].SetValueAs false
        row[6].SetValueAs "Tests if column names with whitespace at end can be unique"
        row
    let thirdRow(range:string) cc =
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 3.
        row[2].SetValueAs "Bye"
        row[3].SetValueAs (System.DateTime(2023,10,16, 20,0,0))
        row[4].SetValueAs true
        row
    let fourthRow(range:string) cc = 
        let row = FsRow(FsRangeAddress(range),cc)
        row[1].SetValueAs 4.269
        row[2].SetValueAs "Outer Space"
        row[3].SetValueAs (System.DateTime(2023,10,17,0,0,0))
        row[4].SetValueAs false
        row

    let rowCollectionA1 = 
        let cells = FsCellsCollection() 
        [|headerRow("A1:F1") ;firstRow("A2:F2") ;secondRow("A3:F3");thirdRow("A4:F4");fourthRow("A5:F5")|] 
        |> Array.map (fun x -> x cells)
    let rowCollectionB4 = 
        let cells = FsCellsCollection() 
        [|headerRow("B4:G4");firstRow("B5:G5");secondRow("B6:G6");thirdRow("B7:G7");fourthRow("B8:G8")|]
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
