namespace FsSpreadsheet.ExcelIO

open DocumentFormat.OpenXml.Spreadsheet
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml


/// Functions for working with SharedStringTables.
module SharedStringTable = 

    /// Functions for working with SharedStringItems.
    module SharedStringItem = 

        /// Gets the string contained in the sharedStringItem.
        let getText (ssi : SharedStringItem) = ssi.InnerText

        /// Sets the string contained in the sharedStringItem.
        let setText text (ssi : SharedStringItem) = 
            ssi.Text <- Text(text)
            ssi

        /// Creates a sharedStringItem containing the given string.
        let create text = 
            new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text))

        /// Adds the sharedStringItem to the sharedStringTable.
        let add (sharedStringItem : SharedStringItem) (sst : SharedStringTable) = 
            sst.Append(sharedStringItem.CloneNode(false) :?> SharedStringItem)
            sst

    /// Creates an empty sharedStringTable.
    let empty() = SharedStringTable() 

    /// Sets an empty sharedStringTable.
    let init (sstPart : SharedStringTablePart) = 
        sstPart.SharedStringTable <- (empty())
        sstPart

    /// Gets the sharedStringTable of the sharedStringTablePart.
    let get (sstPart : SharedStringTablePart) = 
        sstPart.SharedStringTable

    /// Sets the sharedStringtable of the sharedStringTablePart.
    let set (sst:SharedStringTable) (sstPart : SharedStringTablePart) = 
        sstPart.SharedStringTable <- sst
        sstPart

    /// Returns the sharedStringItems contained in the sharedStringTable.
    let toSeq (sst : SharedStringTable) : seq<SharedStringItem> = 
        sst.Elements<SharedStringItem>()
    
    /// Returns the index of the string, or the invers of the element count.
    let getIndexByString (text : string) (sst : SharedStringTable) = 
        let rec loop (en:System.Collections.Generic.IEnumerator<OpenXmlElement>) i =
            match en.MoveNext() with
            | true -> match (en.Current.InnerText = text) with
                      | true -> i
                      | false -> loop en (i+1)
            | false -> ~~~i // invers count
            
        loop (sst.GetEnumerator()) 0

    /// If the string is contained in the sharedStringTable, contains the index of its position.
    let tryGetIndexByString (s : string) (sst : SharedStringTable) = 
        toSeq sst 
        |> Seq.tryFindIndex (fun x -> x.Text.Text = s)

    /// Returns the sharedStringItem at the given index.
    let getText i (sst : SharedStringTable) = 
        toSeq sst
        |> Seq.item i

 

    /// Number of sharedStringItems in the sharedStringTable.
    let count (sst : SharedStringTable) = 
        if sst.Count.HasValue then sst.Count.Value else 0u

    /// Appends the SharedStringItem to the end of the SharedStringTable.
    let append (sharedStringItem:SharedStringItem) (sharedStringTable:SharedStringTable) = 
        sharedStringTable.AppendChild(sharedStringItem) |> ignore
        sharedStringTable.Save()
        sharedStringTable

    /// Inserts text into the sharedStringTable. If the item already exists, returns its index.
    let insertText (text : string) (sharedStringTable : SharedStringTable) = 
        let index = getIndexByString text sharedStringTable
        if index < 0 then 
            let ssi = SharedStringItem.create text
            append ssi sharedStringTable |> ignore
            ~~~index
        else
            index
            
            



