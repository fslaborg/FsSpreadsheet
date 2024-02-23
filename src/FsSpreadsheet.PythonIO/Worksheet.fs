namespace FsSpreadsheet.ExcelPy

module PyWorksheet =

    open Fable.Core
    open FsSpreadsheet
    open Fable.Openpyxl
    open Fable.Core.PyInterop
    
    type WorksheetStatic =
        [<Emit("new $0(parent=$1, title=$2)")>]
        abstract member create: parent:Workbook * title:string -> Worksheet

    [<Import("Worksheet", "openpyxl.worksheet.worksheet")>]
    let Worksheet : WorksheetStatic = nativeOnly

    let fromFsWorksheet (parent : Workbook) (fsWS: FsWorksheet) : Worksheet =
        let pyWS = Worksheet.create(parent,fsWS.Name)
        fsWS.Tables
        |> Seq.iter (fun table -> 
            let pyTable = PyTable.fromFsTable table
            pyWS.add_table(pyTable) |> ignore
        )
        fsWS.CellCollection.GetCells()
        |> Seq.iter (fun cell -> 
            let pyCell = PyCell.fromFsCell cell
            pyWS.cell(cell.Address.RowNumber, cell.Address.ColumnNumber, pyCell) |> ignore
        )
        pyWS


    let toFsWorksheet(pyWS:Worksheet) : FsWorksheet =  
        let fsWS = FsWorksheet(pyWS.title)
        pyWS.tables.values() |> Array.iter (fun table -> 
            let t = PyTable.toFsTable table
            fsWS.AddTable(t) |> ignore
        )
        pyWS.rows |> Array.iteri (fun rowIndex row -> 
            row |> Array.iteri (fun colIndex cell -> 
                let c = PyCell.toFsCell pyWS.title rowIndex colIndex cell
                fsWS.AddCell(c) |> ignore
            )
        )
        fsWS
