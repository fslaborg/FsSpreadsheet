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

    let closedXML() =
        
        // Read xlsx file using closedxml
        let readAssay() = new ClosedXML.Excel.XLWorkbook(assayPath)
        let readStudy() = new ClosedXML.Excel.XLWorkbook(studyPath)
        let readInvestigation() = new ClosedXML.Excel.XLWorkbook(investigationPath)

        readAssay() |> ignore
        readStudy() |> ignore
        readInvestigation() |> ignore
    

    fsSpreadsheet()
    closedXML()

    1