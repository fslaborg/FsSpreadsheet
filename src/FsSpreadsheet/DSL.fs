namespace FsSpreadsheet.DSL

open FsSpreadsheet

type Value = DataType * string

type ColumnIndex = 
    
    | Col of int 

    member self.Index = match self with | Col i -> i

type RowIndex = 
    
    | Row of int

    member self.Index = match self with | Row i -> i

type ColumnElement =
    | IndexedCell of RowIndex * Value
    | UnindexedCell of Value

type RowElement =
    | IndexedCell of ColumnIndex * Value
    | UnindexedCell of Value

type SheetElement = 
    | IndexedRow of RowIndex * RowElement list
    | UnindexedRow of RowElement list
    | IndexedColumn of ColumnIndex * ColumnElement list
    | UnindexedColumn of ColumnElement list

type WorkbookElement =
    | UnnamedSheet of SheetElement list
    | NamedSheet of string * SheetElement list

type Workbook =
    | Workbook of WorkbookElement list

[<AutoOpen>]
type DSL = 

    static member cell (v : obj) = UnindexedCell (DataType.InferCellValue v)
    
    static member cell(columnIndex : ColumnIndex) = fun (v : obj) -> IndexedCell (columnIndex,(DataType.InferCellValue v))

    //static member cell(rowIndex : RowIndex) = fun (v : obj) -> ColumnElement.IndexedCell (rowIndex,(DataType.InferCellValue v))

    static member row (elements : RowElement list) = UnindexedRow elements

    static member row(rowIndex : RowIndex) = fun (elements : RowElement list) -> IndexedRow (rowIndex,elements)

    //static member column (elements : ColumnElement list) = UnindexedColumn elements
   
    static member sheet (elements : SheetElement list) = UnnamedSheet elements

    static member sheet (name : string) = fun (elements : SheetElement list) -> NamedSheet (name, elements)

    static member workbook = fun (elements : WorkbookElement list) -> Workbook elements




type Workbook with
     

    static member internal parseRow (cellCollection : FsCellsCollection) (row : FsRow) (els : RowElement list) =
        let mutable cellIndexSet = 
            els 
            |> List.choose (fun el -> match el with | IndexedCell(i,_) -> Some i.Index | _ -> None)
            |> set
        let getNextIndex () = 
            let mutable i = 1 
            while cellIndexSet.Contains i do
                i <- i + 1
            cellIndexSet <- Set.add i cellIndexSet
            i
        els
        |> List.iter (fun el ->
            match el with 
            | IndexedCell(i,(datatype,value)) -> 
                let cell = row.Cell(i.Index,cellCollection)
                cell.DataType <- datatype
                cell.Value <- value
            | UnindexedCell(datatype,value) -> 
                let cell = row.Cell(getNextIndex(),cellCollection)
                cell.DataType <- datatype
                cell.Value <- value
        )

    static member internal parseSheet (sheet : FsWorksheet) (els : SheetElement list) =
        let mutable rowIndexSet = 
            els 
            |> List.choose (fun el -> match el with | IndexedRow(i,_) -> Some i.Index | _ -> None)
            |> set
        let getNextIndex () = 
            let mutable i = 1 
            while rowIndexSet.Contains i do
                i <- i + 1
            rowIndexSet <- Set.add i rowIndexSet
            i
        els
        |> List.iter (fun el ->
            match el with 
            | IndexedRow(i,rowElements) -> 
                let row = sheet.Row(i.Index)
                Workbook.parseRow sheet.CellCollection row rowElements                
                
            | UnindexedRow(rowElements) -> 
                let row = sheet.Row(getNextIndex())
                Workbook.parseRow sheet.CellCollection row rowElements                   
        )

    member self.Parse() =
        match self with
        | Workbook wbEls -> 
            let workbook = new FsWorkbook()
            wbEls
            |> List.iteri (fun i wbEl ->
                match wbEl with
                | UnnamedSheet sheetEls ->
                    let worksheet = FsWorksheet(sprintf "Sheet%i" (i+1))
                    Workbook.parseSheet worksheet sheetEls
                    workbook.AddWorksheet(worksheet) |> ignore

                | NamedSheet (name,sheetEls) ->
                    let worksheet = FsWorksheet(name)
                    Workbook.parseSheet worksheet sheetEls
                    workbook.AddWorksheet(worksheet) |> ignore            
            )
            workbook