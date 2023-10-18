namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml

module Stylesheet = 
    
    //module Font = 
        
    //    let getDefault() = Font().Color


    let get (doc : SpreadsheetDocument) =
        
        doc.WorkbookPart.WorkbookStylesPart.Stylesheet

    let getOrInit (doc : SpreadsheetDocument) =
        
        match doc.WorkbookPart.WorkbookStylesPart with
        | null -> 
            let ssp = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>()
            ssp.Stylesheet <- new Stylesheet()
            ssp.Stylesheet.CellFormats <- new CellFormats()
            ssp.Stylesheet
        | ssp -> ssp.Stylesheet

    let tryGet (doc : SpreadsheetDocument) =
        match doc.WorkbookPart.WorkbookStylesPart with
        | null -> None
        | ssp -> Some(ssp.Stylesheet)

    module CellFormat = 
        
        let updateCount (stylesheet : Stylesheet) =
            let newCount = stylesheet.CellFormats.Elements<CellFormat>() |> Seq.length
            stylesheet.CellFormats.Count <- UInt32Value(uint32 newCount)

        let count (stylesheet : Stylesheet) =
            if stylesheet.CellFormats = null then 0
            elif stylesheet.CellFormats.Count = null then 0
            else stylesheet.CellFormats.Count.Value |> int

        let getAt (index : int) (stylesheet : Stylesheet) =
            stylesheet.CellFormats.Elements<CellFormat>() |> Seq.item index

        let tryGetAt (index : int) (stylesheet : Stylesheet) =
            stylesheet.CellFormats.Elements<CellFormat>() |> Seq.tryItem index

        let setAt (index : int) (cf : CellFormat) (stylesheet : Stylesheet) =            
            if count stylesheet > index  then
                let previousChild = getAt index stylesheet
                stylesheet.CellFormats.ReplaceChild(cf, previousChild) |> ignore
            if count stylesheet = index then
                stylesheet.CellFormats.AppendChild(cf) |> ignore
            else failwith "Cannot insert style into stylesheet: Index out of range"
            updateCount stylesheet

        let append (cf : CellFormat) (stylesheet : Stylesheet) =
            stylesheet.CellFormats.AppendChild(cf) |> ignore
            updateCount stylesheet