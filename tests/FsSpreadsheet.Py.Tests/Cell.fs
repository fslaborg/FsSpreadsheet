module Cell.Tests

open TestingUtils
open Fable.Openpyxl
open FsSpreadsheet.ExcelPy
open FsSpreadsheet

open Fable.Core

let fromFsCell = testList "fromFsCell" [
    testCase "Boolean" <| fun _ ->
        let fsCell = FsCell(true, DataType.Boolean)
        let pyCell = PyCell.fromFsCell fsCell
        let expected = Some(box true)
        Expect.equal pyCell expected ""
]

//let toFsCell = testList "toFsCell" [
//    testCase "Boolean" <| fun _ ->
//        let pyCell = 

//]


let main = testList "Cell" [
    //toFsCell
    fromFsCell
]
