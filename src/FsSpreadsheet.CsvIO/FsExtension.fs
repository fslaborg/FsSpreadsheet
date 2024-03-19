namespace FsSpreadsheet.CsvIO

open FsSpreadsheet

[<AutoOpen>]
module FsExtensions =

    type FsRow with

        member self.ToSparseRow(cellCollection : FsCellsCollection) : seq<string option> =

            let maxColIndex = self.RangeAddress.LastAddress.ColumnNumber
            let cells = self.Cells
            seq {
                for i = 1 to maxColIndex do
                    cells 
                    |> Seq.tryPick (fun cell ->
                        if cell.ColumnNumber = i then 
                            cell.ValueAsString() |> Some
                        else 
                            None
                    )
            }

    type FsWorksheet with

        member self.ToSparseTable() : seq<seq<string option> option>=
            self.RescanRows()
            self.SortRows()
            let rows = self.Rows
            let maxRowIndex = self.CellCollection.MaxRowNumber
            seq {
                for i = 1 to maxRowIndex do
                    rows 
                    |> Seq.tryPick (fun row ->
                        if row.Index = i then 
                            Some (row.ToSparseRow(self.CellCollection))
                        else 
                            None
                    )
            }

        member self.ToTableString(separator : char option) : string =
            let separator = separator |> Option.defaultValue ','
            self.ToSparseTable()
            |> Seq.map (fun row ->
                match row with
                | None -> ""
                | Some s when Seq.isEmpty s -> ""
                | Some row ->
                    row
                    |> Seq.map (fun cell -> cell |> Option.defaultValue "")
                    |> Seq.reduce (fun rowString cellString -> $"{rowString}{separator}{cellString}")
            )
            |> Seq.reduce (fun tableString rowString ->
                $"{tableString}\n{rowString}"            
            ) 

    type FsWorkbook with

        member self.ToCsvStream(stream : System.IO.MemoryStream,?Separator : char) = 
            let streamWriter = new System.IO.StreamWriter(stream)
            self.GetWorksheets().[0].ToTableString(Separator)
            |> streamWriter.Write
            streamWriter.Flush()

        member self.ToCsvBytes(?Separator : char) =
            use memoryStream = new System.IO.MemoryStream()
            match Separator with
            | Some s -> self.ToCsvStream(memoryStream,s)
            | None -> self.ToCsvStream(memoryStream)
            memoryStream.ToArray()

        member self.ToCsvFile(path,?Separator : char) =
            match Separator with
            | Some s -> self.ToCsvBytes(s)
            | None -> self.ToCsvBytes()
            |> fun bytes -> System.IO.File.WriteAllBytes (path, bytes)

        static member toCsvStream(stream : System.IO.MemoryStream,workbook : FsWorkbook,?Separator : char) =
            match Separator with
            | Some s -> workbook.ToCsvStream(stream,s)
            | None -> workbook.ToCsvStream(stream)
            workbook.ToCsvStream(stream)

        static member toCsvBytes(workbook: FsWorkbook,?Separator : char) =
            match Separator with
            | Some s -> workbook.ToCsvBytes(s)
            | None -> workbook.ToCsvBytes()

        static member toCsvFile(path,workbook: FsWorkbook,?Separator : char) =
            match Separator with
            | Some s -> workbook.ToCsvFile(path,s)
            | None -> workbook.ToCsvFile(path)

type Writer =
    
    static member toCsvStream(stream : System.IO.MemoryStream,workbook : FsWorkbook,?Separator : char) =
        match Separator with
        | Some s -> workbook.ToCsvStream(stream,s)
        | None -> workbook.ToCsvStream(stream)
        workbook.ToCsvStream(stream)

    static member toCsvBytes(workbook: FsWorkbook,?Separator : char) =
        match Separator with
        | Some s -> workbook.ToCsvBytes(s)
        | None -> workbook.ToCsvBytes()

    static member toCsvFile(path,workbook: FsWorkbook,?Separator : char) =
        match Separator with
        | Some s -> workbook.ToCsvFile(path,s)
        | None -> workbook.ToCsvFile(path)