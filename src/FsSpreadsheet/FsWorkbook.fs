namespace FsSpreadsheet


/// <summary>Creates an empty FsWorkbook.</summary>
type FsWorkbook() =
 
    let mutable _worksheets = []

    interface System.IDisposable with
        member self.Dispose() = ()

    // maybe better to leave that to methods...
    //member self.Worksheets 
    //    with get() = _worksheets
    //    and set(value) = _worksheets <- value


    // ------------------
    // NON-STATIC METHODS
    // ------------------
 
    /// Adds an FsWorksheet with given name.
    member self.AddWorksheet(name : string) = 
        let sheet = FsWorksheet name
        _worksheets <- List.append _worksheets [sheet]
        sheet

    /// Adds a given FsWorksheet.
    member self.AddWorksheet(sheet : FsWorksheet) = 
        if _worksheets |> List.exists (fun ws -> ws.Name = sheet.Name) then
            failwithf "Could not add worksheet with name \"%s\" to workbook as it already contains a worksheet with the same name" sheet.Name
        else
            _worksheets <- List.append _worksheets [sheet]
        sheet

    /// Returns all FsWorksheets.
    member self.GetWorksheets() = 
        _worksheets

    /// Removes an FsWorksheet with given name.
    member self.RemoveWorksheet(name : string) =
        let filteredWorksheets =
            match _worksheets |> List.tryFind (fun ws -> ws.Name = name) with
            | Some w -> _worksheets |> List.filter (fun ws -> ws.Name <> name)
            | None -> failwith $"FsWorksheet with name {name} was not found in FsWorkbook."
        _worksheets <- filteredWorksheets
        self

    /// Removes a given FsWorksheet.
    member self.RemoveWorksheet(sheet : FsWorksheet) =
        self.RemoveWorksheet(sheet.Name)
 

    // --------------
    // STATIC METHODS
    // --------------

    /// Adds an FsWorksheet with given name to an FsWorkbook.
    static member addWorksheetWithName (name : string) (workbook : FsWorkbook) = 
        workbook.AddWorksheet name

    /// Adds an FsWorksheet to an FsWorkbook.
    static member addWorksheet (sheet : FsWorksheet) (workbook : FsWorkbook) = 
        workbook.AddWorksheet sheet

    /// Returns all FsWorksheets.
    static member getWorksheets (workbook : FsWorkbook) =
        workbook.GetWorksheets()

    /// Removes an FsWorksheet with given name from an FsWorkbook.
    static member removeWorksheetByName (name : string) (workbook : FsWorkbook) =
        workbook.RemoveWorksheet name

    /// Removes a given FsWorksheet from an FsWorkbook.
    static member removeWorksheet (sheet : FsWorksheet) (workbook : FsWorkbook) =
        workbook.RemoveWorksheet sheet