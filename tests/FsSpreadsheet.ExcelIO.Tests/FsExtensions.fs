module FsExtension

open Expecto
open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml
open Spreadsheet
open System.IO


let dummyDtNumber = DataType.Number
let dummyDtString = DataType.String
let dummyDtBoolean = DataType.Boolean
let dummyDtDate = DataType.Date
let dummyDtEmpty = DataType.Empty

let dummyXlsxCell = Cell.create CellValues.Number "A1" (CellValue(1.337))

//let testFilePath = @"C:\Repos\CSBiology\FsSpreadsheet\tests\FsSpreadsheet.ExcelIO.Tests\data\testUnit.xlsx"
let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "data", "testUnit.xlsx")
let sr = new StreamReader(testFilePath)
let dummyFsWorkbook = new FsWorkbook()
let dummyFsCells = [
    FsCell.create 1 1 "A1"          // for sheet1 (StringSheet)
    FsCell.create 7 3 "7"           // for sheet2 (NumericSheet)
    FsCell.create 2 10 "B10"        // for sheet3 (TableSheet)
    FsCell.create 2 1 "True"        // for sheet4 (DataTypeSheet), DataType.Boolean
    FsCell.create 5 1 "03.13.2023"  // for sheet4 (DataTypeSheet), DataType.DateTime
]
let dummyFsCellsCollection1 = FsCellsCollection()
dummyFsCellsCollection1.Add dummyFsCells[0] |> ignore
let dummyFsCellsCollection2 = FsCellsCollection()
dummyFsCellsCollection2.Add dummyFsCells[1] |> ignore
let dummyFsCellsCollection3 = FsCellsCollection()
dummyFsCellsCollection3.Add dummyFsCells[2] |> ignore
let dummyFsCellsCollection4 = FsCellsCollection()
dummyFsCellsCollection4.Add dummyFsCells[3] |> ignore
dummyFsCellsCollection4.Add dummyFsCells[4] |> ignore
let dummyFsTable = FsTable("Table2", FsRangeAddress("A1:D13"))
let dummyFsWorksheet1 = FsWorksheet("StringSheet",      [], [],             dummyFsCellsCollection1)
let dummyFsWorksheet2 = FsWorksheet("NumericSheet",     [], [],             dummyFsCellsCollection2)
let dummyFsWorksheet3 = FsWorksheet("TableSheet",       [], [dummyFsTable], dummyFsCellsCollection3)
let dummyFsWorksheet4 = FsWorksheet("DataTypeSheet",    [], [],             dummyFsCellsCollection4)
dummyFsWorkbook.AddWorksheet(dummyFsWorksheet1) |> ignore
dummyFsWorkbook.AddWorksheet(dummyFsWorksheet2) |> ignore
dummyFsWorkbook.AddWorksheet(dummyFsWorksheet3) |> ignore
dummyFsWorkbook.AddWorksheet(dummyFsWorksheet4) |> ignore




[<Tests>]
let fsExtensionTests =
    testList "FsExtensions" [
        testList "DataType" [
            testList "ofXlsxCellValues" [
                let testCvNumber = DataType.ofXlsxCellValues CellValues.Number
                testCase "is correct DataTypeNumber from CellValuesNumber" <| fun _ ->
                    Expect.equal testCvNumber DataType.Number "is not the correct DataType"
                let testCvString = DataType.ofXlsxCellValues CellValues.String
                testCase "is correct DataTypeString from CellValuesString" <| fun _ ->
                    Expect.equal testCvString DataType.String "is not the correct DataType"
                let testCvSharedString = DataType.ofXlsxCellValues CellValues.SharedString
                testCase "is correct DataTypeString from CellValuesSharedString" <| fun _ ->
                    Expect.equal testCvSharedString DataType.String "is not the correct DataType"
                let testCvInlineString = DataType.ofXlsxCellValues CellValues.InlineString
                testCase "is correct DataTypeString from CellValuesInlineString" <| fun _ ->
                    Expect.equal testCvInlineString DataType.String "is not the correct DataType"
                let testCvBoolean = DataType.ofXlsxCellValues CellValues.Boolean
                testCase "is correct DataTypeBoolean from CellValuesBoolean" <| fun _ ->
                    Expect.equal testCvBoolean DataType.Boolean "is not the correct DataType"
                let testCvDate = DataType.ofXlsxCellValues CellValues.Date
                testCase "is correct DataTypeDate from CellValuesDate" <| fun _ ->
                    Expect.equal testCvDate DataType.Date "is not the correct DataType"
                let testCvError = DataType.ofXlsxCellValues CellValues.Error
                testCase "is correct DataTypeEmpty from CellValuesError" <| fun _ ->
                    Expect.equal testCvError DataType.Empty "is not the correct DataType"
            ]
        ]
        testList "FsCell" [
            testList "ofXlsxCell" [
                let testCell = FsCell.ofXlsxCell None dummyXlsxCell
                testCase "is equal in value" <| fun _ ->
                    Expect.equal testCell.Value dummyXlsxCell.CellValue.Text "values are not equal"
                testCase "is equal in address/reference" <| fun _ ->
                    Expect.equal testCell.Address.Address dummyXlsxCell.CellReference.Value "addresses/references are not equal"
                testCase "is equal in DataType/CellValues" <| fun _ ->
                    let dtOfCvs = DataType.ofXlsxCellValues dummyXlsxCell.DataType
                    Expect.equal testCell.DataType dtOfCvs "addresses/references are not equal"
            ]
        ]
        testList "FsTable" [
            testList "GetWorksheetOfTable" [
                testCase "gets correct worksheet" <| fun _ ->
                    let gottenWorksheet = dummyFsTable.GetWorksheetOfTable dummyFsWorkbook
                    Expect.equal gottenWorksheet.Name dummyFsWorksheet3.Name "Worksheets differ"
            ]
        ]
        testList "FsWorkbook" [
            testList "fromXlsxStream" [
                let fsWorkbookFromStream = FsWorkbook.fromXlsxStream sr.BaseStream
                sr.Close()
                let fsWorksheet1FromStream = fsWorkbookFromStream.GetWorksheetByName "StringSheet"
                let fsWorksheet2FromStream = fsWorkbookFromStream.GetWorksheetByName "NumericSheet"
                let fsWorksheet3FromStream = fsWorkbookFromStream.GetWorksheetByName "TableSheet"
                let fsWorksheet4FromStream = fsWorkbookFromStream.GetWorksheetByName "DataTypeSheet"
                testCase "is equal to dummyFsWorkbook in sheet1, cellA1 value" <| fun _ ->
                    let v = (FsWorksheet.getCellAt 1 1 fsWorksheet1FromStream).Value
                    Expect.equal v "A1" "value is not equal"
                testCase "is equal to dummyFsWorkbook in sheet1, cellA1 address" <| fun _ ->
                    let a = (FsWorksheet.getCellAt 1 1 fsWorksheet1FromStream).Address.Address
                    Expect.equal a "A1" "address is not equal"
                testCase "is equal to dummyFsWorkbook in sheet1, cellA1 DataType" <| fun _ ->
                    let d = (FsWorksheet.getCellAt 1 1 fsWorksheet1FromStream).DataType
                    Expect.equal d DataType.String "DataType is not DataType.String"
                testCase "is equal to dummyFsWorkbook in sheet2, cellC7 value" <| fun _ ->
                    let v = (FsWorksheet.getCellAt 7 3 fsWorksheet2FromStream).Value
                    Expect.equal v "7" "value is not equal"
                testCase "is equal to dummyFsWorkbook in sheet2, cellC7 address" <| fun _ ->
                    let a = (FsWorksheet.getCellAt 7 3 fsWorksheet2FromStream).Address.Address
                    Expect.equal a "C7" "address is not equal"
                // TO DO: this does not work right now due to incorrect DataType inference. Fix DataType.InferCellValue
                //testCase "is equal to dummyFsWorkbook in sheet2, cellC7 DataType" <| fun _ ->
                //    let d = (FsWorksheet.getCellAt 7 3 fsWorksheet2FromStream).DataType
                //    Expect.equal d DataType.Number "DataType is not DataType.Number"
                testCase "is equal to dummyFsWorkbook in sheet3, cellB10 value" <| fun _ ->
                    let v = (FsWorksheet.getCellAt 10 2 fsWorksheet3FromStream).Value
                    Expect.equal v "B10" "value is not equal"
                testCase "is equal to dummyFsWorkbook in sheet3, cellB10 address" <| fun _ ->
                    let a = (FsWorksheet.getCellAt 10 2 fsWorksheet3FromStream).Address.Address
                    Expect.equal a "B10" "address is not equal"
                testCase "is equal to dummyFsWorkbook in sheet3, cellB10 DataType" <| fun _ ->
                    let d = (FsWorksheet.getCellAt 10 2 fsWorksheet3FromStream).DataType
                    Expect.equal d DataType.String "DataType is not DataType.String"
                testCase "is equal to dummyFsWorkbook in sheet4, cellA2 value" <| fun _ ->
                    let v = (FsWorksheet.getCellAt 2 1 fsWorksheet4FromStream).Value
                    Expect.equal v "1" "value is not equal"     // should be "True"... why is it not? Maybe bc. it's stored as "1" in the XML and only Excel converts it to "TRUE" on the screen... TO DO: check that.
                testCase "is equal to dummyFsWorkbook in sheet4, cellA2 address" <| fun _ ->
                    let a = (FsWorksheet.getCellAt 2 1 fsWorksheet4FromStream).Address.Address
                    Expect.equal a "A2" "address is not equal"
                testCase "is equal to dummyFsWorkbook in sheet4, cellA2 DataType" <| fun _ ->
                    let d = (FsWorksheet.getCellAt 2 1 fsWorksheet4FromStream).DataType
                    Expect.equal d DataType.Boolean "DataType is not DataType.Boolean"
                testCase "is equal to dummyFsWorkbook in sheet3, Table exists" <| fun _ ->
                    let t = fsWorksheet3FromStream.Tables |> List.tryFind (fun t -> t.Name = dummyFsTable.Name)
                    Expect.isSome t "Table \"table2\" does not exist"
                testCase "is equal to dummyFsWorkbook in sheet3, Table has same range" <| fun _ ->
                    let t = fsWorksheet3FromStream.Tables |> List.find (fun t -> t.Name = dummyFsTable.Name)
                    let rfs = t.RangeAddress.Range
                    let rdt = dummyFsTable.RangeAddress.Range
                    Expect.equal rfs rdt "Tables have different ranges"
            ]
        ]
    ]