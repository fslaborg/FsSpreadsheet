module Stylesheet

open TestingUtils
open FsSpreadsheet.Net
open DocumentFormat.OpenXml.Spreadsheet

let private tests_NumberingFormat = testList "NumberingFormat" [
    testList "isDateTime" [
        let testFormat (input: bool*string) =
            let expected, format = input
            testCase format <| fun _ ->
                let numberingFormat = new NumberingFormat()
                numberingFormat.FormatCode <- format
                let isDateTime = Stylesheet.NumberingFormat.isDateTime numberingFormat
                Expect.equal isDateTime expected format
        let formats = [|
            false, "General"
            false, "aaaa"
            true, "dd/mm/yyyy"
            true, "d/m/yy\ h:mm;@"
            true, "m/d/yyyy"
            false, "0.00"
        |]
        for format in formats do 
            yield testFormat format
    ]
]

let main = testList "Stylesheet" [
    tests_NumberingFormat
]