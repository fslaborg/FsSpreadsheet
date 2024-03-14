namespace FsSpreadsheet.Net

open DocumentFormat.OpenXml.Spreadsheet

/// Functions for working with CellData.
module CellData =


    // Formula error (byte)
    type CellError =
        // #NULL!
      | Null = 0x00
        // #DIV/0!
      | DIV0 = 0x07 
        // #VALUE!
      | VALUE = 0x0F
        // #REF!
      | REF = 0x17
        // #NAME?
      | NAME = 0x1D
        // #NUM!
      | NUM = 0x24
        // #N/A
      | NA = 0x2A
        // #GETTING_DATA
      | GettingDATA = 0x2B


    type CellDataValue =
        | Number  of string
        | String  of string
        | Boolean of string
        | Date    of string
        | Error   of string

    /// Creates a CellDataValue from a sharedStringTable and a cell.
    let ofCell (sharedStringTable : SharedStringTable) (cell : Cell) =
        if not (isNull cell.DataType) then  
            let index = int cell.InnerText
            match sharedStringTable |> Seq.tryItem index with 
            | Some value -> 
              match cell.DataType.Value with
              | CellValues.Number  ->  CellDataValue.Number value.InnerText
              | CellValues.Boolean ->  CellDataValue.Boolean value.InnerText
              | CellValues.Date    ->  CellDataValue.Date value.InnerText
              | CellValues.Error   ->  CellDataValue.Error value.InnerText
              | _                  ->  CellDataValue.String value.InnerText
            | None ->
              let value = cell.InnerText
              CellDataValue.Number value
        else
            let value = cell.InnerText
            CellDataValue.Number value      

