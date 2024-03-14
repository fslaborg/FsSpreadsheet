module Table.Tests

open TestingUtils
open Fable.Openpyxl
open FsSpreadsheet.Py
open FsSpreadsheet

open Fable.Core.PyInterop
open Fable.Core

let fromFsTable = testList "fromFsTable" [
    testCase "Table" <| fun _ ->
        let rangeString = "A1:E5"
        let fsTable = FsTable("MyTable", FsRangeAddress(rangeString))
        let pyTable = PyTable.fromFsTable fsTable
        let expected = Table.create("MyTable", rangeString)

        Expect.equal pyTable.name expected.name ""
        Expect.equal pyTable.displayName expected.displayName ""
        Expect.equal pyTable?ref expected?ref ""
        Expect.equal pyTable?ref rangeString ""
]

let toFsTable = testList "toFsTable" [
    testCase "Table" <| fun _ ->
        let pyTable = Table.create("MyTable", "A1:E5")
        let fsTable = PyTable.toFsTable pyTable
        let expected = FsTable("MyTable", FsRangeAddress("A1:E5"))
        Expect.equal fsTable.Name expected.Name ""
        Expect.equal fsTable.RangeAddress.Range expected.RangeAddress.Range ""

]

let main = testList "Table" [
    fromFsTable
    toFsTable
]