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

## Usage_IO

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

Xlsx.toFile(newPath,wb)
```

### Python

```python
from fsspreadsheet.xlsx import Xlsx

path = "path/to/spreadsheet.xlsx"

wb = Xlsx.fromXlsxFile(path)

newPath = "path/to/new/spreadsheet.xlsx"

Xlsx.to_file(newPath,wb)
```
