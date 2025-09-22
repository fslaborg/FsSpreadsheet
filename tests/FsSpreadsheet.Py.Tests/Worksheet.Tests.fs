module Worksheet.Tests

open TestingUtils
open FsSpreadsheet.Py
open FsSpreadsheet
open Fable.Openpyxl
open Fable.Core.PyInterop
open Fable.Core

let wsName = "MySheet"

let fromFsWorksheet = testList "fromFsWorksheet" [
    testCase "Empty" <| fun _ ->
        let wb = Workbook.create()
        let fsWS = FsWorksheet(wsName)
        let pyWS = PyWorksheet.fromFsWorksheet wb fsWS

        Expect.equal pyWS.title wsName "Name did not match"
        Expect.equal pyWS.values.Length 0 "Values did not match" 
    testCase "Simple" <| fun _ ->
        let wb = Workbook.create()
        let fsWS = FsWorksheet(wsName)
        let cell = FsCell("Hello", address = FsAddress.fromString("A1"))
        fsWS.AddCell(cell) |> ignore
        let pyWS = PyWorksheet.fromFsWorksheet wb fsWS

        let expectedCells = [| [| box "Hello" |] |]

        Expect.equal pyWS.title wsName "Name did not match"
        Expect.equal pyWS.values expectedCells "Values did not match"
]


let toFsWorksheet = testList "toFsWorksheet" [
    testCase "Empty" <| fun _ ->
        let wb = Workbook.create()
        let pyWS = wb.create_sheet(wsName)
        let fsWS = PyWorksheet.toFsWorksheet pyWS
        Expect.equal fsWS.Name wsName "Name did not match"
        Expect.equal (fsWS.CellCollection.GetCells() |> Seq.length) 0 "Cells did not match"

]

let main = testList "Worksheet" [
    fromFsWorksheet
    toFsWorksheet
]