module Resources

open FsSpreadsheet

let dummyFsCells =
    [
        [
            FsCell.createWithDataType DataType.String 1 1 "A1"
            FsCell.createWithDataType DataType.String 1 2 "B1"
            FsCell.createWithDataType DataType.String 1 3 "C1"
        ]
        [
            FsCell.createWithDataType DataType.String 2 1 "A2"
            FsCell.createWithDataType DataType.String 2 2 "B2"
            FsCell.createWithDataType DataType.String 2 3 "C2"
        ]
        [
            FsCell.createWithDataType DataType.String 3 1 "A3"
            FsCell.createWithDataType DataType.String 3 2 "B3"
            FsCell.createWithDataType DataType.String 3 3 "C3"
        ]
    ]