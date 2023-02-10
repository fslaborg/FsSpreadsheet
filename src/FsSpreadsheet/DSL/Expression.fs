namespace FsSpreadsheet.DSL

open Microsoft.FSharp.Linq.RuntimeHelpers

module Expression =

    [<NoComparison; NoEquality; Sealed>]
    type OptionalSource<'T>(s : 'T) =
        member this.Source = s

    [<NoComparison; NoEquality; Sealed>]
    type RequiredSource<'T>(s : 'T) =
        member this.Source = s

    [<NoComparison; NoEquality; Sealed>]
    type ExpressionSource<'T>(s : 'T) =

        member this.Source = s

    let eval<'T> q = LeafExpressionConverter.EvaluateQuotation q :?> 'T
