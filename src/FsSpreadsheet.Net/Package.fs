namespace FsSpreadsheet.Net

open DocumentFormat.OpenXml
open System.IO

module Package =

    let tryGetApplication (package : Packaging.Package) =
        let uri = new System.Uri("/docProps/app.xml", System.UriKind.Relative);
        if package.PartExists(uri) then
            let part = package.GetPart(uri)
            use stream = part.GetStream()
            use reader = System.Xml.XmlReader.Create(stream)
            let ns = System.Xml.Linq.XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
            let root = System.Xml.Linq.XElement.Load(reader)
            let app = root.Element(ns + "Application")
            if app <> null then
                Some app.Value
            else None
        else 
            None

    let fixLibrePackage (package : Packaging.Package) =

        let uri = new System.Uri("/xl/webextensions/taskpanes.xml", System.UriKind.Relative);

        package.DeletePart(uri)
        package.CreatePart(uri,contentType = "application/vnd.ms-office.webextensiontaskpanes+xml")
        |> ignore


    let isLibrePackage (package : Packaging.Package) =
        match tryGetApplication package with
        | Some app -> app.Contains "LibreOffice"
        | None -> false