module TestObjects

open FsSpreadsheet

let sheet1Name = "MySheet1"
let sheet2Name = "MySheet2"

let sheet1() =
    let ws = new FsWorksheet(sheet1Name)
    [
        FsCell.createWithDataType DataType.String 1 1 "A1"
        FsCell.createWithDataType DataType.String 1 2 "B1"
        FsCell.createWithDataType DataType.String 1 3 "C1"

        FsCell.createWithDataType DataType.String 2 1 "A2"
        FsCell.createWithDataType DataType.String 2 2 "B2"
        FsCell.createWithDataType DataType.String 2 3 "C2"

        FsCell.createWithDataType DataType.String 3 1 "A3"
        FsCell.createWithDataType DataType.String 3 2 "B3"
        FsCell.createWithDataType DataType.String 3 3 "C3"        
    ]
    |> List.iter (fun c -> ws.Row(c.RowNumber).[c.ColumnNumber].SetValueAs c.Value)
    ws

let sheet2() =
    let ws = new FsWorksheet(sheet2Name)
    [
        FsCell.createWithDataType DataType.Number 1 1 1.
        FsCell.createWithDataType DataType.Number 1 2 2.
        FsCell.createWithDataType DataType.Number 1 3 3.
        FsCell.createWithDataType DataType.Number 1 4 4.

        FsCell.createWithDataType DataType.Number 2 1 5.
        FsCell.createWithDataType DataType.Number 2 2 6.
        FsCell.createWithDataType DataType.Number 2 3 7.
        FsCell.createWithDataType DataType.Number 2 4 8.    
    ]
    |> List.iter (fun c -> ws.Row(c.RowNumber).[c.ColumnNumber].SetValueAs c.Value)
    ws