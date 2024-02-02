module FsSpreadsheet.ExcelIO.ZipArchiveReader


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
    let string = "4"
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

    static member dateTimeFormats = Set.ofList [14..22] 
    static member customFormats = Set.ofList [164 .. 180] 

    static member isDateTime (styles : Styles) (cf : CellFormat) =
            // if numberformatid is between 14 and 18 it is standard date time format.
            // custom formats are given in the range of 164 to 180, all none default date time formats fall in there.
            if Set.contains cf.NumberFormatId CellFormat.dateTimeFormats then 
                true
            elif Set.contains cf.NumberFormatId CellFormat.customFormats then   
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


let parseRelationsships (relationships : ZipArchiveEntry) : Relationships =
    try 
        let relationshipsStream = relationships.Open()
        let relationshipsReader = System.Xml.XmlReader.Create(relationshipsStream)
        let mutable rels = [||]
        while relationshipsReader.Read() do
            if relationshipsReader.NodeType = XmlNodeType.Element && relationshipsReader.Name = "Relationship" then
                let id = relationshipsReader.GetAttribute("Id")
                let typ = relationshipsReader.GetAttribute("Type")
                let target = 
                    let t = relationshipsReader.GetAttribute("Target")
                    if t.StartsWith "xl/" then t
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
        let sharedStringsStream = sharedStrings.Open()
        let sharedStringsReader = System.Xml.XmlReader.Create(sharedStringsStream)
        [|
        while sharedStringsReader.Read() do
            if sharedStringsReader.NodeType = System.Xml.XmlNodeType.Element && sharedStringsReader.Name = "si" then
                sharedStringsReader.ReadToDescendant("t") |> ignore
                yield sharedStringsReader.ReadElementContentAsString()
        |]
    with
    | _ -> [||]

let getStyles (wb : WorkBook) =
    let styles = wb.GetEntry("xl/styles.xml")
    let stylesStream = styles.Open()
    let stylesReader = System.Xml.XmlReader.Create(stylesStream)
    let mutable numberFormats = Map.empty
    let mutable cellFormats = Array.empty
    while stylesReader.Read() do
        if stylesReader.NodeType = XmlNodeType.Element && stylesReader.Name = "numFmts" then
            let subReader = stylesReader.ReadSubtree()
            let numFmts = 
                [|
                    while subReader.Read() do
                        if subReader.NodeType = XmlNodeType.Element && subReader.Name = "numFmt" then
                            let id = subReader.GetAttribute("numFmtId") |> int                    
                            let formatCode = subReader.GetAttribute("formatCode")
                            yield id, {Id = id; FormatCode = formatCode}
                |]
            numberFormats <- Map.ofArray numFmts
        if stylesReader.NodeType = XmlNodeType.Element && stylesReader.Name = "cellXfs" then
            let subReader = stylesReader.ReadSubtree()
            let cellFmts = 
                [|
                    while subReader.Read() do
                        if subReader.NodeType = XmlNodeType.Element && subReader.Name = "xf" then
                            let numFmtId = subReader.GetAttribute("numFmtId") |> int
                            let applyNumberFormat = subReader.GetAttribute("applyNumberFormat") |> bool
                            yield {NumberFormatId = numFmtId; ApllyNumberFormat = applyNumberFormat}
                |]
            cellFormats <- cellFmts
    {NumberFormats = numberFormats; CellFormats = cellFormats}

       
let parseTable (sheet : ZipArchiveEntry) =
    try
        let stream = sheet.Open()
        let reader = System.Xml.XmlReader.Create(stream)
        while reader.Name <> "table" do
            reader.Read() |> ignore 
        let area = reader.GetAttribute("ref")
        let ra = FsRangeAddress(area)    
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
        FsTable(name, ra, totalsRowShown, true)
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

let parseDataType (styles : Styles) (dataType : string) (style : string) (formula : string) = 
    if formula <> null then
        // LibreOffice annotates boolean values as formulas instead of boolean datatypes
        if formula = "TRUE()" || formula = "FALSE()" then
            DataType.Boolean
        else
            DataType.Number
    //https://stackoverflow.com/a/13178043/12858021
    //https://stackoverflow.com/a/55425719/12858021
    // if styleindex is not null and datatype is null we propably have a DateTime field.
    // if datatype would not be null it could also be boolean, as far as i tested it ~Kevin F 13.10.2023
    elif style <> null && (dataType = null || dataType = DType.number) then
        try
            let cellFormat : CellFormat = styles.CellFormats.[int style]
            if (*cellFormat <> null &&*) CellFormat.isDateTime styles cellFormat then
                DataType.Date
                        
            else
                DataType.Number
        with
        | _ -> DataType.Number
    else 
        match dataType with
        | DType.number -> DataType.Number
        | DType.boolean -> DataType.Boolean
        | DType.date -> DataType.Date
        | DType.error -> DataType.Empty
        | DType.inlineString
        | DType.sharedString
        | DType.string -> DataType.String
        | _ -> DataType.Number

let parseCell (sst : SharedStrings) (styles : Styles) (value : string) (dataType : string) (style : string) (formula : string) : obj*DataType =

    let cellValueString = if dataType <> null && dataType = DType.sharedString then sst.[int value] else value
    let dt = parseDataType styles dataType style formula
    match dt with
    | Date ->
        try 
            // datetime is written as float counting days since 1900. 
            // We use the .NET helper because we really do not want to deal with datetime issues.
            System.DateTime.FromOADate(float cellValueString)
        with 
        | _ -> cellValueString
    | Boolean ->
        // boolean is written as int/float either 0 or null
        match cellValueString.ToLower() with 
        | "1" | "true" -> true 
        | "0" | "false" -> false 
        | _ -> cellValueString
    | Number -> 
        try 
            float cellValueString
        with 
            | _ -> 
                cellValueString
    | Empty | String -> cellValueString
    ,
    dt

let parseWorksheet (name : string) (styles : Styles) (sharedStrings : SharedStrings) (sheet : ZipArchiveEntry) (wb : WorkBook) =
    try 
        let ws = FsWorksheet(name)
        let stream = sheet.Open()
        let reader = System.Xml.XmlReader.Create(stream)
        let relationships = getWsRelationships sheet.FullName wb
        while reader.Read() do
            if reader.NodeType = XmlNodeType.Element && reader.Name = "c" then
                let r = reader.GetAttribute("r")
                let t = reader.GetAttribute("t")
                let s = reader.GetAttribute("s")
            
                let mutable v = null
                let mutable f = null
                let cellReader = reader.ReadSubtree()
                while cellReader.Read() do
                    if cellReader.NodeType = XmlNodeType.Element && cellReader.Name = "v" then
                        v <- cellReader.ReadElementContentAsString()
                    if cellReader.NodeType = XmlNodeType.Element && cellReader.Name = "f" then
                        f <- cellReader.ReadElementContentAsString()
                if v <> null && v <> "" then
                    let cellValue,dataType = parseCell sharedStrings styles v t s f
                    let cell = FsCell(cellValue,dataType = dataType,address = FsAddress(r))
                    ws.AddCell(cell) |> ignore
            if reader.NodeType = XmlNodeType.Element && reader.Name = "tablePart" then
                printfn "TablePart"
                let id = reader.GetAttribute("r:id")
                printfn "ID: %s" id
                let rel = relationships.[id]
                printf "Target: %s" rel.Target
                let table = 
                    wb.GetEntry(rel.Target)
                    |> parseTable
                printfn "tableref: %s" table.RangeAddress.Range
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
    let wbStream = wbPart.Open()
    let wbReader = System.Xml.XmlReader.Create(wbStream)
    while wbReader.Read() do
        if wbReader.Name = "sheet" then
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

    let fromStream (stream : Stream) = 
        use zip = new ZipArchive(stream)
        fromZipArchive zip

    let fromBytes (bytes : byte []) = 
        use ms = new MemoryStream(bytes)
        fromStream ms

    let fromFile (path : string) =
       use fs = File.OpenRead(path)
       fromStream fs