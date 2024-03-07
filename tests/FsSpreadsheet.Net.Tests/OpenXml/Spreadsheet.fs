module Spreadsheet

open Expecto
open FsSpreadsheet.Net
open DocumentFormat.OpenXml


let testFilePath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "../data", "testUnit.xlsx")
// *fox = from OpenXml, to distinguish between objects from FsSpreadsheet.Net
let ssdFox = Packaging.SpreadsheetDocument.Open(testFilePath, false)
let wbpFox = ssdFox.WorkbookPart
let sstpFox = wbpFox.SharedStringTablePart
let sstFox = sstpFox.SharedStringTable
let sstFoxInnerText = sstFox.InnerText
let wsp1Fox = (wbpFox.WorksheetParts |> Array.ofSeq)[0]
let cbsi1Fox =      // get the Cells, but with their real values (inferred from the SST) not their SST index
    wsp1Fox.Worksheet.Descendants<Spreadsheet.Cell>() 
    |> Array.ofSeq
    |> Array.map (
        fun c ->
            if c.DataType <> null && c.DataType.Value = Spreadsheet.CellValues.SharedString then
                let index = int c.CellValue.InnerText
                let item = sstFox.Elements<OpenXmlElement>() |> Seq.item index
                let value = item.InnerText
                c.CellValue.Text <- value
                c
            else
                c
    )


//let testSsdFox2 = Packaging.SpreadsheetDocument.Open(testFilePath, false)

//let testDoc = Spreadsheet.fromFile testFilePath false

//testSsdFox = testDoc
//testSsdFox = testSsdFox2



[<Tests>]
let spreadsheetTests =
    testList "Spreadsheet" [
        let ssd = Spreadsheet.fromFile testFilePath false
        // is it even possible to test this?
        //testList "fromFile" [
            //testCase "is equal to testSsd" <| fun _ ->
            //    Expect.equal testDoc ssdFox "Both testFiles differ"
            //    testSsd.Close()
        //]
        testList "tryGetSharedStringTable" [
            let sst = Spreadsheet.tryGetSharedStringTable ssdFox
            testCase "is Some" <| fun _ ->
                Expect.isSome sst "Is None"
            testCase "is equal to testSstFox" <| fun _ ->
                Expect.equal sst.Value sstFox "Differs"
            testCase "inner text is correct" <| fun _ ->
                Expect.equal sst.Value.InnerText sstFoxInnerText "Differs"
        ]
        testList "getWorkbookPart" [
            let wbp = Spreadsheet.getWorkbookPart ssdFox
            testCase "is equal to wbpFox" <| fun _ ->
                Expect.equal wbp wbpFox "Differs"
        ]
        testList "getCellsBySheetIndex" [
            let cbsi1 = Spreadsheet.getCellsBySheetIndex 1u ssd false |> Array.ofSeq
            // not applicable since Cell arrays and Cells always differ from each other (even if you compare the same cell, e.g. `cell1 = cell1`)
            //testCase "is equal to cbsi1Fox" <| fun _ ->
            //    Expect.equal cbsi1 cbsi1Fox "Differs"
            // therefore, test individual Cell properties:
            testCase "element 0 is equal in CellReference to element 0 from OpenXML" <| fun _ ->
                Expect.equal cbsi1[0].CellReference cbsi1Fox[0].CellReference "Differs"
            testCase "element 10 is equal in CellReference to element 10 from OpenXML" <| fun _ ->
                Expect.equal cbsi1[10].CellReference cbsi1Fox[10].CellReference "Differs"
            testCase "element 20 is equal in CellReference to element 20 from OpenXML" <| fun _ ->
                Expect.equal cbsi1[20].CellReference cbsi1Fox[20].CellReference "Differs"
            testCase "element 0 is equal in CellValueText to element 0 from OpenXML" <| fun _ ->
                Expect.equal cbsi1[0].CellValue.Text cbsi1Fox[0].CellValue.Text "Differs"
            testCase "element 10 is equal in CellValueText to element 10 from OpenXML" <| fun _ ->
                Expect.equal cbsi1[10].CellValue.Text cbsi1Fox[10].CellValue.Text "Differs"
            testCase "element 20 is equal in CellValueText to element 20 from OpenXML" <| fun _ ->
                Expect.equal cbsi1[20].CellValue.Text cbsi1Fox[20].CellValue.Text "Differs"
        ]
    ]









//testSsdFox.Close()