namespace FsSpreadsheet.Py

module PyTable =

    open Fable.Core
    open FsSpreadsheet
    open Fable.Openpyxl
    open Fable.Core.PyInterop
    

    type tablestyle =
        abstract member name: unit -> string
        abstract member showFirstColumn: unit -> bool
        abstract member showLastColumn: unit -> bool
        abstract member showRowStripes: unit -> bool
        abstract member showColumnStripes: unit -> bool

    type TableStyleStatic =
        [<Emit("new $0(name=$1, showFirstColumn=$2, showLastColumn=$3, showRowStripes=$4, showColumnStripes=$5)")>]
        abstract member create: name:string * showFirstColumn:bool * showLastColumn:bool * showRowStripes:bool * showColumnStripes:bool -> tablestyle
    
    [<Import("TableStyleInfo", "openpyxl.worksheet.table")>]
    let TableStyleInfo : TableStyleStatic = nativeOnly

    let defaultTableStyle() = TableStyleInfo.create("TableStyleMedium9", false, false, true, true)

    let fromFsTable (fsTable: FsTable) : Table =
        let table = Table.create(fsTable.Name,fsTable.RangeAddress.Range)
        table?tableStyleInfo <- defaultTableStyle()
        table

    let toFsTable(table:Table) =
        let name = if isNull table.displayName then table.name else table.displayName
        let ref = table.ref
        FsTable(name,FsRangeAddress(ref))
