# FsSpreadsheet
Spreadsheet creation and manipulation in FSharp

<table>
  <thead>
    <tr>
      <th>Latest Release</th>
      <th>Downloads</th>
      <th>Target</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>
        <a href="https://pypi.org/project/fsspreadsheet/">
          <img src="https://img.shields.io/pypi/v/fsspreadsheet?logo=pypi" alt="latest release" />
        </a>
      </td>
      <td>
        <a href="https://pepy.tech/project/siren-dsl">
          <img alt="Pepy Total Downlods" src="https://img.shields.io/pepy/dt/siren-dsl?label=fsspreadsheet&color=blue" />
        </a>
      </td>
      <td>Python</td>
    </tr>
    <!-- js package -->
    <tr>
      <td>
        <a href="https://www.npmjs.com/package/@fslab/fsspreadsheet">
          <img src="https://img.shields.io/npm/v/@fslab/fsspreadsheet?logo=npm" alt="latest release" />
        </a>
      </td>
      <td>
        <a href="https://www.npmjs.com/package/@fslab/fsspreadsheet">
          <img src="https://img.shields.io/npm/dt/@fslab/fsspreadsheet?label=@fslab/fsspreadsheet" alt="downloads" />
        </a>
      </td>
      <td>JavaScript</td>
    </tr>
    <!-- f# nuget package core -->
    <tr>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet/">
          <img src="https://img.shields.io/nuget/v/FsSpreadsheet?logo=nuget" alt="latest release" />
        </a>
      </td>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet/">
          <img src="https://img.shields.io/nuget/dt/FsSpreadsheet?label=FsSpreadsheet" alt="downloads" />
        </a>
      </td>
      <td></td>
    </tr>
    <!-- f# nuget package net -->
    <tr>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet.Net/">
          <img src="https://img.shields.io/nuget/v/FsSpreadsheet.Net?logo=nuget" alt="latest release" />
        </a>
      </td>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet.Net/">
          <img src="https://img.shields.io/nuget/dt/FsSpreadsheet.Net?label=FsSpreadsheet.Net" alt="downloads" />
        </a>
      </td>
      <td>.NET</td>
    </tr>
    <!-- f# nuget package js -->
    <tr>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet.Js/">
          <img src="https://img.shields.io/nuget/v/FsSpreadsheet.Js?logo=nuget" alt="latest release" />
        </a>
      </td>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet.Js/">
          <img src="https://img.shields.io/nuget/dt/FsSpreadsheet.Js?label=FsSpreadsheet.Js" alt="downloads" />
        </a>
      </td>
      <td>Fable JavaScript</td>
    </tr>
    <!-- f# nuget package py -->
    <tr>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet.Py/">
          <img src="https://img.shields.io/nuget/v/FsSpreadsheet.Py?logo=nuget" alt="latest release" />
        </a>
      </td>
      <td>
        <a href="https://www.nuget.org/packages/FsSpreadsheet.Py/">
          <img src="https://img.shields.io/nuget/dt/FsSpreadsheet.Py?label=FsSpreadsheet.Py" alt="downloads" />
        </a>
      </td>
      <td>Fable Python</td>
    </tr>
  </tbody>

</table>

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

### Contribution guidelines

- Make sure that all contributions run on all three languages: F#, Javascript and Python
- Please add failing tests prior to fixing a bug against which to code
  - If applicable, include issue number in test name as such: `"worksOnFilledTable (issue #100)"`
- Make use of [XML tags](https://github.com/fslaborg/FsSpreadsheet/issues/10) to comment your code as such:
   ```fsharp
   /// <summary>
   /// Checks if there is an FsCell at given column index of a given FsRow.
   /// </summary>
   /// <param name="colIndex">The number of the column where the presence of an FsCell shall be checked.</param>
   /// <param name="row"></param>
   static member hasCellAt colIndex (row : FsRow) =
	  row.HasCellAt colIndex
   ```
