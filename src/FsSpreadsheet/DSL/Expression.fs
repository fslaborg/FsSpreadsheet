namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Linq.RuntimeHelpers

module Expression =

    let eval<'T> q = LeafExpressionConverter.EvaluateQuotation q :?> 'T
