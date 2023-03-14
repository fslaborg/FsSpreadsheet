module FsExtension

open Expecto
open FsSpreadsheet
open FsSpreadsheet.ExcelIO
open DocumentFormat.OpenXml
open System.IO


let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "data", "testUnit.xlsx")
let sr = new StreamReader(testFilePath)
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.ExcelIO
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let sstpFox = wbpFox.SharedStringTablePart
let sstFox = sstpFox.SharedStringTable
let sstFoxInnerText = sstFox.InnerText
let wsp1Fox = (wbpFox.WorksheetParts |> Array.ofSeq)[0]
let cbsi1Fox = wsp1Fox.Worksheet.Descendants<Spreadsheet.Cell>() |> Array.ofSeq
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




[<Tests>]
let fsExtensionTests =
    testList "FsExtensions" [
        testList "FsWorkbook" [
            testList "FromXlsxStream" [
                let fsWorkbookFromStream = FsWorkbook.fromXlsxStream sr.BaseStream
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
                testCase "is equal to dummyFsWorkbook in sheet2, cellC7 DataType" <| fun _ ->
                    let d = (FsWorksheet.getCellAt 7 3 fsWorksheet2FromStream).DataType
                    Expect.equal d DataType.Number "DataType is not DataType.Number"
                testCase "is equal to dummyFsWorkbook in sheet2, cellC7 value" <| fun _ ->
                    let v = (FsWorksheet.getCellAt 2 10 fsWorksheet3FromStream).Value
                    Expect.equal v "B10" "value is not equal"
                testCase "is equal to dummyFsWorkbook in sheet3, cellB10 address" <| fun _ ->
                    let a = (FsWorksheet.getCellAt 2 10 fsWorksheet3FromStream).Address.Address
                    Expect.equal a "B10" "address is not equal"
                testCase "is equal to dummyFsWorkbook in sheet3, cellB10 DataType" <| fun _ ->
                    let d = (FsWorksheet.getCellAt 2 10 fsWorksheet3FromStream).DataType
                    Expect.equal d DataType.String "DataType is not DataType.String"
                testCase "is equal to dummyFsWorkbook in sheet3, cellB10 value" <| fun _ ->
                    let v = (FsWorksheet.getCellAt 2 1 fsWorksheet4FromStream).Value
                    Expect.equal v "True" "value is not equal"
                testCase "is equal to dummyFsWorkbook in sheet2, cellC7 address" <| fun _ ->
                    let a = (FsWorksheet.getCellAt 2 1 fsWorksheet4FromStream).Address.Address
                    Expect.equal a "A2" "address is not equal"
                testCase "is equal to dummyFsWorkbook in sheet2, cellC7 DataType" <| fun _ ->
                    let d = (FsWorksheet.getCellAt 2 1 fsWorksheet4FromStream).DataType
                    Expect.equal d DataType.Boolean "DataType is not DataType.Boolean"
            ]
        ]
    ]