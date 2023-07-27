# FsSpreadsheet
Spreadsheet creation and manipulation in FSharp

## DSL 
```fsharp
#r "nuget: FsSpreadsheet"

open FsSpreadsheet.DSL

let dslTree = 

    workbook {
        sheet "MySheet" {
            row {
                cell {1}
                cell {2}
                cell {3}
            }
            row {
                4
                5
                6
            }
        }
    }


let spreadsheet = dslTree.Value.Parse()
```
## ExcelIO

```fsharp

#r "nuget: FsSpreadsheet.ExcelIO"

open FsSpreadsheet.ExcelIO

spreadsheet.ToFile(excelFilePath)

```

------->

![image](https://user-images.githubusercontent.com/17880410/167841765-f67e1fa2-3806-4f32-9223-bdecc8253568.png)

## Code Examples

```fsharp
let parseWorkbook wb =
    let tables = getAnnotationTables wb
    let worksheets = wb.GetWorksheets()
    // get worksheet and its AnnotationTable as tuple
    let worksheetsAndTables =
        tables
        |> List.map (
            fun t ->
                let associatedWs = 
                    worksheets
                    |> List.find (
                        fun ws -> 
                            ws.Tables
                            |> List.exists (fun t2 -> t2.Name = t.Name)
                    )
                associatedWs, t
        )
    worksheetsAndTables
    |> List.map (
        fun (ws,t) ->
            let headerCells = getHeaderCellsOf ws.CellCollection t |> List.ofSeq
            let ioColumns = 
                headerCells
                |> List.map (parseIOColumns ws.CellCollection t)
            let termRelatedBuildingBlocks =
                headerCells
                |> groupTermRelatedBuildingBlocks t ws.CellCollection
                |> List.map (parseTermRelatedBuildingBlocks true t ws.CellCollection)
            ws.Name,
            {|
                IOColumns = ioColumns
                TermRelatedBuildingBlocks = termRelatedBuildingBlocks
            |}
    )
```


## Develop

### Build QuickStart

If not already done,
1. install .NET SDK
2. install Node.js

In any shell, run
1. `dotnet tool restore`
4. `npm install`
5. `build.cmd <target>` where `<target>` may be
    - if `<target>` is empty, it just runs dotnet build after cleaning everything
    - `runtests` to run unit tests
      - `runtestsjs` to only run JS unit tests
	  - `runtestsdotnet` to only run .NET unit tests
    - `releasenotes semver:<version>` where `<version>` may be `major`, `minor`, or `patch` to update RELEASE_NOTES.md
    - `pack` to create a NuGet release
      - `packprelease` to create a NuGet prerelease
    - `builddocs` to create docs
      - `builddocsprerelease` to create prerelease docs
  	- `watchdocs` to create docs and run them locally
  	- `watchdocsprelease` to create prerelease docs and run them locally
    - `publishnuget` to create a NuGet release and publish it
      - `publishnugetprelease` to create a NuGet prerelease and publish it
