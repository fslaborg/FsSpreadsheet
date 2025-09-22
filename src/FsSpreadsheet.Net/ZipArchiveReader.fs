module FsSpreadsheet.Net.ZipArchiveReader

open System.IO.Compression
open System.Xml

open FsSpreadsheet

let inline bool (v : string) = 
    match v with
    | null | "0" -> false
    | "1" -> true
    | _ -> 
          let v' = v.ToLower()
          if v' = "true" then true
            elif v' = "false" then false
            else failwith "Invalid boolean value"


module DType = 
    [<Literal>]
    let boolean = "b"
    [<Literal>]
    let number = "n"
    [<Literal>]
    let error = "e"
    [<Literal>]
    let sharedString = "s"
    [<Literal>]
    let string = "str"
    [<Literal>]
    let inlineString = "5"
    [<Literal>]
    let date = "d"

type Relationship = 
    { Id : string
      Type : string
      Target : string }

type WorkBook = ZipArchive

type SharedStrings = string []

type Relationships = Map<string,Relationship>

let dateTimeFormats = Set.ofList [14..22] 
let customFormats = Set.ofList [164 .. 180] 


type NumberFormat =
    { Id : int
      FormatCode : string }

    static member isDateTime (numberFormat : NumberFormat) =
        let format = numberFormat.FormatCode
        let input = System.DateTime.Now.ToString(format, System.Globalization.CultureInfo.InvariantCulture)
        let dt = System.DateTime.ParseExact(
            input, 
            format, 
            System.Globalization.CultureInfo.InvariantCulture, 
            System.Globalization.DateTimeStyles.NoCurrentDateDefault
        )
        dt <> Unchecked.defaultof<System.DateTime>

and CellFormat =
    {
        NumberFormatId : int 
        ApllyNumberFormat : bool
    }

    static member isDateTime (styles : Styles) (cf : CellFormat) =
            // if numberformatid is between 14 and 18 it is standard date time format.
            // custom formats are given in the range of 164 to 180, all none default date time formats fall in there.
            if Set.contains cf.NumberFormatId dateTimeFormats then 
                true
            elif Set.contains cf.NumberFormatId customFormats then   
                styles.NumberFormats.TryFind cf.NumberFormatId
                |> Option.map NumberFormat.isDateTime
                |> Option.defaultValue true             
            else 
                false

and Styles = 
    {
        NumberFormats : Map<int,NumberFormat>
        CellFormats : CellFormat []
    }


module XmlReader =
    let isElemWithName (reader : System.Xml.XmlReader) (name : string) =
        reader.NodeType = XmlNodeType.Element && (reader.Name = name || reader.Name = "x:" + name)

let parseRelationsships (relationships : ZipArchiveEntry) : Relationships =
    try 
        use relationshipsStream = relationships.Open()
        use relationshipsReader = System.Xml.XmlReader.Create(relationshipsStream)
        let mutable rels = [||]
        while relationshipsReader.Read() do
            if XmlReader.isElemWithName relationshipsReader "Relationship" then
                let id = relationshipsReader.GetAttribute("Id")
                let typ = relationshipsReader.GetAttribute("Type")
                let target = 
                    let t = relationshipsReader.GetAttribute("Target")
                    if t.StartsWith "xl/" then t
                    elif t.StartsWith "/xl/" then t.Replace("/xl/","xl/")
                    elif t.StartsWith "../" then t.Replace("../","xl/")
                    else "xl/" + t
                rels <- Array.append rels [|{Id = id; Type = typ; Target = target}|]
        rels
        |> Array.map (fun r -> r.Id, r) |> Map.ofArray
    with
    | _ -> Map.empty

let getWbRelationships (wb : WorkBook) =
    wb.GetEntry("xl/_rels/workbook.xml.rels")
    |> parseRelationsships

let getWsRelationships (ws : string) (wb : WorkBook) =
    wb.GetEntry(ws.Replace("worksheets/","worksheets/_rels/").Replace(".xml",".xml.rels"))
    |> parseRelationsships

let getSharedStrings (wb : WorkBook) : SharedStrings =
    try

        let sharedStrings = wb.GetEntry("xl/sharedStrings.xml")
        use sharedStringsStream = sharedStrings.Open()
        use sharedStringsReader = System.Xml.XmlReader.Create(sharedStringsStream)
        [|
        while sharedStringsReader.Read() do
            if XmlReader.isElemWithName sharedStringsReader "si" then
                use subReader = sharedStringsReader.ReadSubtree()
                while subReader.Read() do
                    if XmlReader.isElemWithName subReader "t" then
                        yield subReader.ReadElementContentAsString()
        |]
    with
    | _ -> [||]

let getStyles (wb : WorkBook) =
    let styles = wb.GetEntry("xl/styles.xml")
    use stylesStream = styles.Open()
    use stylesReader = System.Xml.XmlReader.Create(stylesStream)
    let mutable numberFormats = Map.empty
    let mutable cellFormats = Array.empty
    while stylesReader.Read() do
        if XmlReader.isElemWithName stylesReader "numFmts" then
            use subReader = stylesReader.ReadSubtree()
            let numFmts = 
                [|
                    while subReader.Read() do
                        if XmlReader.isElemWithName subReader "numFmt" then
                            let id = subReader.GetAttribute("numFmtId") |> int                    
                            let formatCode = subReader.GetAttribute("formatCode")
                            yield id, {Id = id; FormatCode = formatCode}
                |]
            numberFormats <- Map.ofArray numFmts
        if XmlReader.isElemWithName stylesReader "cellXfs" then
            use subReader = stylesReader.ReadSubtree()
            let cellFmts = 
                [|
                    while subReader.Read() do
                        if XmlReader.isElemWithName subReader "xf" then
                            let numFmtId = subReader.GetAttribute("numFmtId") |> int
                            let applyNumberFormat = subReader.GetAttribute("applyNumberFormat") |> bool
                            yield {NumberFormatId = numFmtId; ApllyNumberFormat = applyNumberFormat}
                |]
            cellFormats <- cellFmts
    {NumberFormats = numberFormats; CellFormats = cellFormats}

       
let parseTable (sheet : ZipArchiveEntry) =
    try
        use stream = sheet.Open()
        use reader = System.Xml.XmlReader.Create(stream)
        let mutable t = None
        while reader.Read() do
            if XmlReader.isElemWithName reader "table" then
                let area = reader.GetAttribute("ref")
                let ra = FsRangeAddress.fromString(area)    
                let totalsRowShown = 
                    let attribute = reader.GetAttribute("totalsRowShown")
                    match attribute with
                    | null 
                    | "0" -> false
                    | "1" -> true
                    | _ -> false
                let name =
                    let dn = reader.GetAttribute("displayName")
                    if dn = null then 
                       reader.GetAttribute("name")
                    else dn
                t <- Some (FsTable(name, ra, totalsRowShown, true))
        if t.IsNone then
            failwith "No table found"
        else
            t.Value
    with
    | err -> failwithf "Error while parsing table \"%s\":%s" sheet.FullName err.Message

//zip.Entries
//|> Seq.map (fun e -> e.Name)
//|> Seq.toArray

// Apply the parseTable function to every zip entry starting with "xl/tables"
let getTables (wb : WorkBook) =
    wb.Entries
    |> Seq.choose (fun e -> 
        if e.FullName.StartsWith("xl/tables") && e.FullName <> "xl/tables" && e.FullName <> "xl/tables/"   then 
            parseTable e
            |> Some 
        else None)
    |> Seq.toArray

let parseCell (sst : SharedStrings) (styles : Styles) (value : string) (dataType : string) (style : string) (formula : string) : obj*DataType = 
    // LibreOffice annotates boolean values as formulas instead of boolean datatypes
    if formula <> null && formula = "TRUE()"  then
        true,DataType.Boolean
    elif formula <> null && formula = "FALSE()" then
        false,DataType.Boolean
    else
        let cellValueString = if dataType <> null && dataType = DType.sharedString then sst.[int value] else value
        //https://stackoverflow.com/a/13178043/12858021
        //https://stackoverflow.com/a/55425719/12858021
        // if styleindex is not null and datatype is null we propably have a DateTime field.
        // if datatype would not be null it could also be boolean, as far as i tested it ~Kevin F 13.10.2023
        if style <> null && (dataType = null || dataType = DType.number) then
            try
                let cellFormat : CellFormat = styles.CellFormats.[int style]
                if (*cellFormat <> null &&*) CellFormat.isDateTime styles cellFormat then
                    System.DateTime.FromOADate(float cellValueString), DataType.Date                      
                else
                    float value, DataType.Number
            with
            | _ -> value, DataType.Number
        else 
            match dataType with
            | DType.boolean -> 
                match cellValueString.ToLower() with 
                | "1" | "true" -> true 
                | "0" | "false" -> false 
                | _ -> cellValueString
                ,
                DataType.Boolean
            | DType.date -> 
                try 
                    // datetime is written as float counting days since 1900. 
                    // We use the .NET helper because we really do not want to deal with datetime issues.
                    System.DateTime.FromOADate(float cellValueString), DataType.Date
                with 
                | _ -> cellValueString, DataType.Date
            | DType.error -> cellValueString, DataType.Empty
            | DType.inlineString
            | DType.sharedString
            | DType.string -> cellValueString, DataType.String
            | _ -> 
                try 
                    float cellValueString
                with 
                | _ -> cellValueString
                ,
                DataType.Number

let parseWorksheet (name : string) (styles : Styles) (sharedStrings : SharedStrings) (sheet : ZipArchiveEntry) (wb : WorkBook) =
    try 
        let ws = FsWorksheet(name)
        let relationships = getWsRelationships sheet.FullName wb
        use stream = sheet.Open()
        use reader = System.Xml.XmlReader.Create(stream)
        while reader.Read() do
            if XmlReader.isElemWithName reader "c" then
                let r = reader.GetAttribute("r")
                let t = reader.GetAttribute("t")
                let s = reader.GetAttribute("s")
            
                let mutable v = null
                let mutable f = null
                use cellReader = reader.ReadSubtree()
                while cellReader.Read() do
                    if XmlReader.isElemWithName cellReader "v" then
                        v <- cellReader.ReadElementContentAsString()
                    if XmlReader.isElemWithName cellReader "f" then
                        f <- cellReader.ReadElementContentAsString()
                if v <> null && v <> "" || f <> null then
                    let cellValue,dataType = parseCell sharedStrings styles v t s f
                    let cell = FsCell(cellValue,dataType = dataType,address = FsAddress.fromString(r))
                    ws.AddCell(cell) |> ignore
            if XmlReader.isElemWithName reader "tablePart" then
                let id = reader.GetAttribute("r:id")
                let rel = relationships.[id]
                let table = 
                    wb.GetEntry(rel.Target)
                    |> parseTable
                ws.AddTable(table) |> ignore
        reader.Close()
        ws
    with
    | err -> failwithf "Error while parsing worksheet \"%s\":%s" name err.Message

let parseWorkbook (wb : ZipArchive) =
    let newWb = new FsWorkbook()
    let styles = getStyles wb
    let sst = getSharedStrings wb
    //let tables = getTables wb
    let relationships = getWbRelationships wb
    let wbPart = wb.GetEntry("xl/workbook.xml")
    use wbStream = wbPart.Open()
    use wbReader = System.Xml.XmlReader.Create(wbStream)
    while wbReader.Read() do
        if XmlReader.isElemWithName wbReader "sheet" then
            let name = wbReader.GetAttribute("name")
            let id = wbReader.GetAttribute("r:id")
            let rel = relationships.[id]
            let sheet = wb.GetEntry(rel.Target)
            let ws = parseWorksheet name styles sst sheet wb
            ws.RescanRows()
            newWb.AddWorksheet (ws) |> ignore
    newWb


module FsWorkbook = 
    
    open System.IO

    let fromZipArchive (wb : ZipArchive) = 
        parseWorkbook wb

    let fromXlsxStream (stream : Stream) = 
        use zip = new ZipArchive(stream)
        fromZipArchive zip

    let fromXlsxBytes (bytes : byte []) = 
        use ms = new MemoryStream(bytes)
        fromXlsxStream ms

    let fromXlsxFile (path : string) =
       use fs = File.OpenRead(path)
       fromXlsxStream fs