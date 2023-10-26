namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml

module Stylesheet = 
    
    module Font = 
        
        let getDefault() = 

            Font(
                FontSize = FontSize(Val = DoubleValue(11.)),
                Color = Color(Theme = UInt32Value(uint32 1)),
                FontName = FontName(Val = StringValue("Calibri")),
                FontFamilyNumbering = FontFamilyNumbering(Val = Int32Value(int32 2)),
                FontScheme = FontScheme(Val = EnumValue(FontSchemeValues.Minor))
            )         

        let updateCount (stylesheet : Stylesheet) =
            let newCount = stylesheet.Fonts.Elements<CellFormat>() |> Seq.length
            stylesheet.Fonts.Count <- UInt32Value(uint32 newCount)

        let initDefaultFonts() =
            let f = Fonts(Count = UInt32Value(1ul))
            f.AppendChild(getDefault()) |> ignore
            f

    module Fill = 
        
        let getDefault() = 
            Fill(
                PatternFill = PatternFill(PatternType = EnumValue(PatternValues.None))
            )

        let updateCount (stylesheet : Stylesheet) =
            let newCount = stylesheet.Fills.Elements<CellFormat>() |> Seq.length
            stylesheet.Fills.Count <- UInt32Value(uint32 newCount)

        let initDefaultFills() =
            let f = Fills(Count = UInt32Value(1ul))
            f.AppendChild(getDefault()) |> ignore
            f

    module Border = 
        
        let getDefault() = 
            Border(
                LeftBorder = LeftBorder(),
                RightBorder = RightBorder(),
                TopBorder = TopBorder(),
                BottomBorder = BottomBorder(),
                DiagonalBorder = DiagonalBorder()
            )

        let updateCount (stylesheet : Stylesheet) =
            let newCount = stylesheet.Borders.Elements<CellFormat>() |> Seq.length
            stylesheet.Borders.Count <- UInt32Value(uint32 newCount)

        let initDefaultBorders() =
            let f = Borders(Count = UInt32Value(1ul))
            f.AppendChild(getDefault()) |> ignore
            f

    module NumberingFormat =

        let get (id : int) (stylesheet : Stylesheet) = 
            stylesheet.NumberingFormats.Elements<NumberingFormat>()
            |> Seq.find (fun nf -> nf.NumberFormatId.Value = uint32 id)

        let tryGet (id : int) (stylesheet : Stylesheet) = 
            try 
                get id stylesheet
                |> Some 
            with 
            | _ -> None

        let getFormatCode (nf : NumberingFormat) = 
            nf.FormatCode.Value

        // Libre does set default numbers to "General" custom format code, so we need to check for that.
        // Floats for example are set to "0.00", so we need a whitelist for isDateTime instead of a blacklist for any else.
        // https://stackoverflow.com/a/72012646/12858021
        let isDateTime (nf : NumberingFormat) = 
            let format = getFormatCode nf 
            let input = System.DateTime.Now.ToString(format, System.Globalization.CultureInfo.InvariantCulture)
            let dt = System.DateTime.ParseExact(
                input, 
                format, 
                System.Globalization.CultureInfo.InvariantCulture, 
                System.Globalization.DateTimeStyles.NoCurrentDateDefault
            )
            dt <> Unchecked.defaultof<System.DateTime>

    module CellFormat = 
        
        let isDateTime (stylesheet : Stylesheet) (cf : CellFormat) =
            // if numberformatid is between 14 and 18 it is standard date time format.
            // custom formats are given in the range of 164 to 180, all none default date time formats fall in there.
            let dateTimeFormats = [14..22] |> List.map (uint32 >> UInt32Value)
            let customFormats = [164 .. 180] |> List.map (uint32 >> UInt32Value)
            if List.contains cf.NumberFormatId dateTimeFormats then 
                true
            elif List.contains cf.NumberFormatId customFormats then
                NumberingFormat.tryGet (cf.NumberFormatId.Value |> int) stylesheet
                |> Option.map NumberingFormat.isDateTime
                |> Option.defaultValue true             
            else 
                false

        let structurallyEquals (cf1 : CellFormat) (cf2 : CellFormat) = 
            cf1.BorderId = cf2.BorderId
            && cf1.FillId = cf2.FillId
            && cf1.FontId = cf2.FontId
            && cf1.NumberFormatId = cf2.NumberFormatId
            && cf1.ApplyNumberFormat = cf2.ApplyNumberFormat

        let updateCount (stylesheet : Stylesheet) =
            let newCount = stylesheet.CellFormats.Elements<CellFormat>() |> Seq.length
            stylesheet.CellFormats.Count <- UInt32Value(uint32 newCount)

        let count (stylesheet : Stylesheet) =
            if stylesheet.CellFormats = null then 0
            elif stylesheet.CellFormats.Count = null then 0
            else stylesheet.CellFormats.Count.Value |> int

        let tryGetIndex (cellFormat : CellFormat) (stylesheet : Stylesheet) =
            if stylesheet.CellFormats = null then None
            else 
                stylesheet.CellFormats.Elements<CellFormat>()
                |> Seq.tryFindIndex (structurallyEquals cellFormat)

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
            
        let appendOrGetIndex (cf : CellFormat) (stylesheet : Stylesheet) =
            match tryGetIndex cf stylesheet with
            | Some i -> i
            | None ->
                append cf stylesheet
                updateCount stylesheet
                (count stylesheet) - 1

        let getDefault () =
            CellFormat(
                NumberFormatId = UInt32Value(0ul),
                FontId = UInt32Value(0ul),
                FillId = UInt32Value(0ul),
                BorderId = UInt32Value(0ul)
                //FormatId = UInt32Value(0ul)                                   
            )

        let getDefaultDate () =
            CellFormat(
                NumberFormatId = UInt32Value(14ul),
                FontId = UInt32Value(0ul),
                FillId = UInt32Value(0ul),
                BorderId = UInt32Value(0ul),
                //FormatId = UInt32Value(0ul),
                ApplyNumberFormat = BooleanValue(true)
            )

        let getDefaultDateTime () =
            CellFormat(
                NumberFormatId = UInt32Value(22ul),
                FontId = UInt32Value(0ul),
                FillId = UInt32Value(0ul),
                BorderId = UInt32Value(0ul),
                //FormatId = UInt32Value(0ul),
                ApplyNumberFormat = BooleanValue(true)                                  
            )

        let initDefaultCellFormats() =
            let f = CellFormats(Count = UInt32Value(1ul))
            f.AppendChild(getDefault()) |> ignore
            f

    let get (doc : SpreadsheetDocument) =
        
        doc.WorkbookPart.WorkbookStylesPart.Stylesheet

    let getOrInit (doc : SpreadsheetDocument) =
        
        match doc.WorkbookPart.WorkbookStylesPart with
        | null -> 
            let ssp = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>()
            ssp.Stylesheet <- new Stylesheet()
            ssp.Stylesheet.CellFormats <- CellFormat.initDefaultCellFormats()
            ssp.Stylesheet.Borders <- Border.initDefaultBorders()
            ssp.Stylesheet.Fills <- Fill.initDefaultFills()
            ssp.Stylesheet.Fonts <- Font.initDefaultFonts()
            ssp.Stylesheet
        | ssp -> ssp.Stylesheet

    let tryGet (doc : SpreadsheetDocument) =
        match doc.WorkbookPart.WorkbookStylesPart with
        | null -> None
        | ssp -> Some(ssp.Stylesheet)