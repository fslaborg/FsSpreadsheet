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

#r "nuget: FsSpreadsheet.Net"

open FsSpreadsheet.Net

spreadsheet.ToFile(excelFilePath)

```

------->

![image](https://user-images.githubusercontent.com/17880410/167841765-f67e1fa2-3806-4f32-9223-bdecc8253568.png)

## Code Examples

```fsharp
let tables = workbook.GetTables()
let worksheets = workbook.GetWorksheets()
// get worksheet and its table as tuple
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
```


## Development

### Requirements

- [nodejs and npm](https://nodejs.org/en/download)
    - verify with `node --version` (Tested with v18.16.1)
    - verify with `npm --version` (Tested with v9.2.0)
- [.NET SDK](https://dotnet.microsoft.com/en-us/download)
    - verify with `dotnet --version` (Tested with 7.0.306)
- [Python](https://www.python.org/downloads/)
    - verify with `py --version` (Tested with 3.12.2)

### Local Setup

1. Setup dotnet tools

   `dotnet tool restore`

2. Install NPM dependencies
   
   `npm install`

3. Setup python environment
    
   `py -m venv .venv`

4. Install [Poetry](https://python-poetry.org/) and dependencies

   1. `.\.venv\Scripts\python.exe -m pip install -U pip setuptools`
   2. `.\.venv\Scripts\python.exe -m pip install poetry`
   3. `.\.venv\Scripts\python.exe -m poetry install --no-root`

Verify correct setup with `./build.cmd runtests` 

5. `build.cmd <target>` where `<target>` may be
    - if `<target>` is empty, it just runs dotnet build after cleaning everything
    - `runtests` to run unit tests
      - `runtestsjs` to only run JS unit tests
	  - `runtestsdotnet` to only run .NET unit tests
      - `runtestpy` to only run Python unit tests
    - `releasenotes semver:<version>` where `<version>` may be `major`, `minor`, or `patch` to update RELEASE_NOTES.md
    - `pack` to create a NuGet release
      - `packprelease` to create a NuGet prerelease
    - `builddocs` to create docs
      - `builddocsprerelease` to create prerelease docs
  	- `watchdocs` to create docs and run them locally
  	- `watchdocsprelease` to create prerelease docs and run them locally
    - `release` to create a NuGet, NPM, PyPI and GitHub release 
