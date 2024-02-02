open FsSpreadsheet
open FsSpreadsheet.ExcelIO

open ClosedXML.Excel

[<EntryPoint>] 
let main argv =
    
    let assayPath = @"C:\Users\HLWei\Downloads\ipk-gabi-wheat\assays\alpha_lattice\isa.assay.xlsx"
    let studyPath = @"C:\Users\HLWei\Downloads\ipk-gabi-wheat\studies\alpha_lattice_trial_data\isa.study.xlsx"
    let investigationPath = @"C:\Users\HLWei\Downloads\ipk-gabi-wheat\isa.investigation.xlsx"

    let fsSpreadsheet() = 


        let readAssay() = FsWorkbook.fromXlsxFile assayPath
        let readStudy() = FsWorkbook.fromXlsxFile studyPath
        let readInvestigation() = FsWorkbook.fromXlsxFile investigationPath

        readAssay() |> ignore
        readStudy() |> ignore
        readInvestigation() |> ignore

    let zipArchiveReader() = 


        let readAssay() = ZipArchiveReader.FsWorkbook.fromFile assayPath
        let readStudy() = ZipArchiveReader.FsWorkbook.fromFile studyPath
        let readInvestigation() = ZipArchiveReader.FsWorkbook.fromFile investigationPath
        let bigFile() = ZipArchiveReader.FsWorkbook.fromFile @"C:\Users\HLWei\source\repos\IO\FsSpreadsheet\tests\TestUtils\TestFiles\BigFile.xlsx"


        readInvestigation() |> ignore
        readAssay() |> ignore
        readStudy() |> ignore
        bigFile() |> ignore

    let closedXML() =
        
        // Read xlsx file using closedxml
        let readAssay() = new ClosedXML.Excel.XLWorkbook(assayPath)
        let readStudy() = new ClosedXML.Excel.XLWorkbook(studyPath)
        let readInvestigation() = new ClosedXML.Excel.XLWorkbook(investigationPath)

        readAssay() |> ignore
        readStudy() |> ignore
        readInvestigation() |> ignore
    
    let randomReadArchives () =

        let readArchive(p : string) = 
            let zip = System.IO.Compression.ZipFile.OpenRead(p)
            let e1 = zip.GetEntry("xl/worksheets/sheet1.xml")
            use stream1 = e1.Open()
            use reader1 = System.Xml.XmlReader.Create(stream1)
            while reader1.Read() do
                ()
            let e2 = zip.GetEntry("xl/worksheets/sheet2.xml")
            use stream2 = e2.Open()
            use reader2 = System.Xml.XmlReader.Create(stream2)
            while reader2.Read() do
                ()
            let e3 = zip.GetEntry("xl/worksheets/sheet3.xml")
            use stream3 = e3.Open()
            use reader3 = System.Xml.XmlReader.Create(stream3)
            while reader3.Read() do
                ()

        let readAssay() = readArchive(assayPath)
        let readStudy() = readArchive(studyPath)
        let readInvestigation() = readArchive(investigationPath)

        readAssay() |> ignore
        readStudy() |> ignore
        readInvestigation() |> ignore

    closedXML()
    fsSpreadsheet()
    zipArchiveReader()
    //randomReadArchives ()
    1