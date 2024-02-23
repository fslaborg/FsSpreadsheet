namespace FsSpreadsheet.ExcelPy

module PyTable =

    open Fable.Core
    open FsSpreadsheet
    open Fable.Openpyxl
    open Fable.Core.PyInterop
    
    let fromFsTable (fsTable: FsTable) : Table =
        Table.create(fsTable.Name,fsTable.RangeAddress.Range)
        
    let toFsTable(table:Table) =
        let name = if isNull table.displayName then table.name else table.displayName
        let ref = table?ref
        FsTable(name,FsRangeAddress(ref))
