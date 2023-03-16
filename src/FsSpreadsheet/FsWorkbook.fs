﻿namespace FsSpreadsheet


/// <summary>Creates an empty FsWorkbook.</summary>
type FsWorkbook() =
 
    let mutable _worksheets = []

    interface System.IDisposable with
        member self.Dispose() = ()

    // maybe better to leave that to methods...
    //member self.Worksheets 
    //    with get() = _worksheets
    //    and set(value) = _worksheets <- value


    // -------
    // METHODS
    // -------
 
    /// Adds an FsWorksheet with given name.
    member self.AddWorksheet(name : string) = 
        let sheet = FsWorksheet name
        _worksheets <- List.append _worksheets [sheet]
        sheet

    /// Adds an FsWorksheet with given name to an FsWorkbook.
    static member addWorksheetWithName (name : string) (workbook : FsWorkbook) = 
        workbook.AddWorksheet name |> ignore
        workbook

    /// Adds a given FsWorksheet.
    member self.AddWorksheet(sheet : FsWorksheet) = 
        if _worksheets |> List.exists (fun ws -> ws.Name = sheet.Name) then
            failwithf "Could not add worksheet with name \"%s\" to workbook as it already contains a worksheet with the same name" sheet.Name
        else
            _worksheets <- List.append _worksheets [sheet]
        sheet

    /// Adds an FsWorksheet to an FsWorkbook.
    static member addWorksheet (sheet : FsWorksheet) (workbook : FsWorkbook) = 
        workbook.AddWorksheet sheet  |> ignore
        workbook

    /// Returns all FsWorksheets.
    member self.GetWorksheets() = 
        _worksheets

    /// Returns all FsWorksheets.
    static member getWorksheets (workbook : FsWorkbook) =
        workbook.GetWorksheets()

    /// <summary>Returns the FsWorksheet with the given name if it exists in the FsWorkbook. Else returns None.</summary>
    member self.TryGetWorksheetByName(sheetName) =
        _worksheets |> List.tryFind (fun w -> w.Name = sheetName)

    /// <summary>Returns the FsWorksheet with the given name if it exists in a given FsWorkbook. Else returns None.</summary>
    static member tryGetWorksheetByName sheetName (workbook : FsWorkbook) =
        workbook.TryGetWorksheetByName sheetName

    /// <summary>Returns the FsWorksheet with the given name.</summary>
    /// <exception cref="System.Exception">if FsWorksheet with given name is not present in the FsWorkkbook.</exception>
    member self.GetWorksheetByName(sheetName) =
        try (self.TryGetWorksheetByName sheetName).Value
        with _ -> failwith $"FsWorksheet with name {sheetName} is not present in the FsWorkbook."

    /// <summary>Returns the FsWorksheet with the given name from an FsWorkbook.</summary>
    /// <exception cref="System.Exception">if FsWorksheet with given name is not present in the FsWorkkbook.</exception>
    static member getWorksheetByName sheetName (workbook : FsWorkbook) =
        workbook.GetWorksheetByName sheetName

    /// <summary>Removes an FsWorksheet with given name.</summary>
    /// <exception cref="System.Exception">if FsWorksheet with given name is not present in the FsWorkkbook.</exception>
    member self.RemoveWorksheet(name : string) =
        let filteredWorksheets =
            match _worksheets |> List.tryFind (fun ws -> ws.Name = name) with
            | Some w -> _worksheets |> List.filter (fun ws -> ws.Name <> name)
            | None -> failwith $"FsWorksheet with name {name} was not found in FsWorkbook."
        _worksheets <- filteredWorksheets
        self

    /// Removes an FsWorksheet with given name from an FsWorkbook.
    static member removeWorksheetByName (name : string) (workbook : FsWorkbook) =
        workbook.RemoveWorksheet name  |> ignore
        workbook

    /// <summary>Removes a given FsWorksheet.</summary>
    /// <exception cref="System.Exception">if FsWorksheet with given name is not present in the FsWorkkbook.</exception>
    member self.RemoveWorksheet(sheet : FsWorksheet) =
        self.RemoveWorksheet(sheet.Name) |> ignore
        self

    /// Removes a given FsWorksheet from an FsWorkbook.
    static member removeWorksheet (sheet : FsWorksheet) (workbook : FsWorkbook) =
        workbook.RemoveWorksheet sheet  |> ignore
        workbook

    /// <summary>Returns all FsTables from the FsWorkbook.</summary>
    member self.GetTables() =
        self.GetWorksheets()
        |> List.collect (fun s -> s.Tables)

    /// <summary>Returns all FsTables from an FsWorkbook.</summary>
    static member getTables (workbook : FsWorkbook) =
        workbook.GetTables()