module FsCell

open Expecto
open FsSpreadsheet

[<Tests>]
let fsWorksheetTest =
    testList "Worksheet FsCell data" [               
        testList "Data | DataType | Adress" [
            let fscellA1_string  = FsCell.create 1 1 "A1"
            let fscellB1_num     = FsCell.create 1 2 1
            let fscellA2_bool    = FsCell.create 1 2 true
            
            let worksheet = FsWorksheet.

            testCase "DataType string" <| fun _ ->
                Expect.equal fscellA1_string.DataType DataType.String "is not the expected DataType.String"
            

        ]
    ]