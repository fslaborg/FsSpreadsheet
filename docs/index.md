# FsSpreadsheet Documentation

FsSpreadsheet is a library for reading and writing spreadsheets in F#, Javascript and Python using Fable. It is a wrapper of the Python library openpyxl (https://openpyxl.readthedocs.io/en/stable/) and the Javascript library exceljs (https://www.npmjs.com/package/@nfdi4plants/exceljs).

## Installation

### F#

Project
```shell
dotnet add package FsSpreadsheet.Net
```

Script
```fsharp
#r "nuget: FsSpreadsheet.Net"

open FsSpreadsheet
open FsSpreadsheet.Net
```

### Javascript

```shell
npm install @fslab/fsspreadsheet
```

### Python

```shell
pip install fsspreadsheet
```

## Usage_Xlsx_IO

### F#

```fsharp
open FsSpreadsheet
open FsSpreadsheet.Net

let path = "path/to/spreadsheet.xlsx"

let wb = FsWorkbook.fromXlsxFile(path)

let newPath = "path/to/new/spreadsheet.xlsx"

wb.ToXlsxFile(newPath)
```

### Javascript

```javascript
import { Xlsx } from '@fslab/fsspreadsheet/Xlsx.js';

const path = "path/to/spreadsheet.xlsx"

const wb = Xlsx.fromXlsxFile(path)

const newPath = "path/to/new/spreadsheet.xlsx"

Xlsx.toXlsxFile(newPath,wb)
```

### Python

```python
from fsspreadsheet.xlsx import Xlsx

path = "path/to/spreadsheet.xlsx"

wb = Xlsx.from_xlsx_file(path)

newPath = "path/to/new/spreadsheet.xlsx"

Xlsx.to_xlsx_file(newPath,wb)
```


## Usage_Json_IO

### F#

```fsharp
open FsSpreadsheet
open FsSpreadsheet.Net

let path = "path/to/spreadsheet.json"

let wb = FsWorkbook.fromJsonFile(path)

let newPath = "path/to/new/spreadsheet.json"

wb.ToJsonFile(newPath)
```

### Javascript

```javascript
import { Json } from '@fslab/fsspreadsheet/Json.js';

const path = "path/to/spreadsheet.json"

const wb = Json.fromJsonFile(path)

const newPath = "path/to/new/spreadsheet.json"

Json.toJsonFile(newPath,wb)
```

### Python

```python
from fsspreadsheet.json import Json

path = "path/to/spreadsheet.json"

wb = Json.from_json_file(path)

newPath = "path/to/new/spreadsheet.json"

Json.to_json_file(newPath,wb)
```
