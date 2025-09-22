module FsAddress

open Fable.Pyxpecto
open FsSpreadsheet


let testAddress1 = FsAddress.fromString("B5")
let testAddress2 = FsAddress(3, 2)
let testAddress3 = FsAddress(4, 8, true, true)
let testAddress4 = FsAddress.fromString("D2", true, true)
let testAddress5 = FsAddress.fromString("Z69")
let testAddress6 = FsAddress(5, 2)

let main = 
    testList "FsAddress" [
        testList "Constructors" [
            testList "cellAddressString" [
                testCase "RowNumber" <| fun _ ->
                    Expect.equal testAddress1.RowNumber 5 "RowNumber differs"
                testCase "ColumnNumber" <| fun _ ->
                    Expect.equal testAddress1.ColumnNumber 2 "ColumnNumber differs"
                testCase "Address" <| fun _ ->
                    Expect.equal testAddress1.Address "B5" "Excel-style address differs"
                testCase "FixedRow" <| fun _ ->
                    Expect.isFalse testAddress1.FixedRow "FixedRow is not false"
                testCase "FixedColumn" <| fun _ ->
                    Expect.isFalse testAddress1.FixedColumn "FixedColumn is not false"
            ]
            testList "rowNumber, columnNumber" [
                testCase "RowNumber" <| fun _ ->
                    Expect.equal testAddress2.RowNumber 3 "RowNumber differs"
                testCase "ColumnNumber" <| fun _ ->
                    Expect.equal testAddress2.ColumnNumber 2 "ColumnNumber differs"
                testCase "Address" <| fun _ ->
                    Expect.equal testAddress2.Address "B3" "Excel-style address differs"
                testCase "FixedRow" <| fun _ ->
                    Expect.isFalse testAddress2.FixedRow "FixedRow is not false"
                testCase "FixedColumn" <| fun _ ->
                    Expect.isFalse testAddress2.FixedColumn "FixedColumn is not false"
            ]
            testList "rowNumber, columnNumber, fixedRow, fixedColumn" [
                testCase "RowNumber" <| fun _ ->
                    Expect.equal testAddress3.RowNumber 4 "RowNumber differs"
                testCase "ColumnNumber" <| fun _ ->
                    Expect.equal testAddress3.ColumnNumber 8 "ColumnNumber differs"
                testCase "Address" <| fun _ ->
                    Expect.equal testAddress3.Address "H4" "Excel-style address differs"
                testCase "FixedRow" <| fun _ ->
                    Expect.isFalse testAddress3.FixedRow "FixedRow is not false"
                testCase "FixedColumn" <| fun _ ->
                    Expect.isFalse testAddress3.FixedColumn "FixedColumn is not false"
            ]
            testList "rowNumber, columnLetter, fixedRow, fixedColumn" [
                testCase "RowNumber" <| fun _ ->
                    Expect.equal testAddress4.RowNumber 2 "RowNumber differs"
                testCase "ColumnNumber" <| fun _ ->
                    Expect.equal testAddress4.ColumnNumber 4 "ColumnNumber differs"
                testCase "Address" <| fun _ ->
                    Expect.equal testAddress4.Address "D2" "Excel-style address differs"
                testCase "FixedRow" <| fun _ ->
                    Expect.isFalse testAddress4.FixedRow "FixedRow is not false"
                testCase "FixedColumn" <| fun _ ->
                    Expect.isFalse testAddress4.FixedColumn "FixedColumn is not false"
            ]
        ]
        testList "UpdateIndices" [
            testList "rowIndex, colIndex" [     // @Contributors: 1 testList per overload pls
                do testAddress5.UpdateIndices(4,1) |> ignore
                testCase "RowNumber" <| fun _ ->
                    Expect.equal testAddress5.RowNumber 4 "RowNumber differs"
                testCase "ColumnNumber" <| fun _ ->
                    Expect.equal testAddress5.ColumnNumber 1 "ColumnNumber differs"
            ]
        ]
        testList "ToIndices" [
            testList "rowIndex, colIndex" [
                let rowIndex, colIndex = testAddress1.ToIndices()
                testCase "RowNumber" <| fun _ ->
                    Expect.equal rowIndex 5 "RowNumber differs"
                testCase "ColumnNumber" <| fun _ ->
                    Expect.equal colIndex 2 "ColumnNumber differs"
            ]
        ]
        testList "Compare" [
            testCase "testAddress1 vs testAddress2" <| fun _ ->
                let result = testAddress1.Compare testAddress2
                Expect.isFalse result "Addresses do not differ"
            testCase "testAddress1 vs testAddress6" <| fun _ ->
                let result = testAddress1.Compare testAddress6
                Expect.isTrue result "Addresses differ"
        ]
        testList "Equals" [
            testCase "testAddress1 vs testAddress2" <| fun _ ->
                let result = testAddress1.Equals testAddress2
                Expect.isFalse result "Addresses do not differ"
            testCase "testAddress1 vs testAddress6" <| fun _ ->
                let result = testAddress1.Equals testAddress6
                Expect.isTrue result "Addresses differ"
        ]
        testList "GetHashCode" [
            testCase "testAddress1 vs testAddress2" <| fun _ ->
                let hash1 = testAddress1.GetHashCode()
                let hash2 = testAddress2.GetHashCode()
                Expect.isFalse (hash1 = hash2) "Hash codes do not differ"
            testCase "testAddress1 vs testAddress6" <| fun _ ->
                let hash1 = testAddress1.GetHashCode()
                let hash2 = testAddress6.GetHashCode()
                Expect.isTrue (hash1 = hash2) "Hash codes differ"      
        ]
    ]