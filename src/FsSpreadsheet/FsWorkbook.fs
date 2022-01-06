namespace FsSpreadsheet


type FsWorkbook() =
 
    let mutable worksheets = []
 
    member self.AddWorksheet(name : string) = 
        let sheet = FsWorksheet (name)
        worksheets <- List.append worksheets [sheet]
        sheet

    member self.AddWorksheet(sheet : FsWorksheet) = 
        if worksheets |> List.exists (fun ws -> ws.Name = sheet.Name) then
            failwithf "Could not add worksheet with name \"%s\" to workbook as it already contains a worksheet with the same name" sheet.Name
        else
            worksheets <- List.append worksheets [sheet]
        sheet

    member self.GetWorksheets() = worksheets

         
    interface System.IDisposable with
         
        member self.Dispose() = ()
 