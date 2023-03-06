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
                        if cell.WorksheetColumn = i then 
                            Option.Some cell.Value
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

        member self.ToStream(stream : System.IO.MemoryStream,?Separator : char) = 
            let streamWriter = new System.IO.StreamWriter(stream)
            self.GetWorksheets().Head.ToTableString(Separator)
            |> streamWriter.Write
            streamWriter.Flush()

        member self.ToBytes(?Separator : char) =
            use memoryStream = new System.IO.MemoryStream()
            match Separator with
            | Some s -> self.ToStream(memoryStream,s)
            | None -> self.ToStream(memoryStream)
            memoryStream.ToArray()

        member self.ToFile(path,?Separator : char) =
            match Separator with
            | Some s -> self.ToBytes(s)
            | None -> self.ToBytes()
            |> fun bytes -> System.IO.File.WriteAllBytes (path, bytes)

        static member toStream(stream : System.IO.MemoryStream,workbook : FsWorkbook,?Separator : char) =
            match Separator with
            | Some s -> workbook.ToStream(stream,s)
            | None -> workbook.ToStream(stream)
            workbook.ToStream(stream)

        static member toBytes(workbook: FsWorkbook,?Separator : char) =
            match Separator with
            | Some s -> workbook.ToBytes(s)
            | None -> workbook.ToBytes()

        static member toFile(path,workbook: FsWorkbook,?Separator : char) =
            match Separator with
            | Some s -> workbook.ToFile(path,s)
            | None -> workbook.ToFile(path)

type Writer =
    
    static member toStream(stream : System.IO.MemoryStream,workbook : FsWorkbook,?Separator : char) =
        match Separator with
        | Some s -> workbook.ToStream(stream,s)
        | None -> workbook.ToStream(stream)
        workbook.ToStream(stream)

    static member toBytes(workbook: FsWorkbook,?Separator : char) =
        match Separator with
        | Some s -> workbook.ToBytes(s)
        | None -> workbook.ToBytes()

    static member toFile(path,workbook: FsWorkbook,?Separator : char) =
        match Separator with
        | Some s -> workbook.ToFile(path,s)
        | None -> workbook.ToFile(path)