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

