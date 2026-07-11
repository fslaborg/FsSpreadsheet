module FsSpreadsheet.Js.Tests

open Fable.Core
open Fable.Core.JsInterop
open Fable.Pyxpecto
open TestingUtils
open FsSpreadsheet.CrossAsync

#if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
[<Emit("process.on('unhandledRejection', function (reason) { $0(reason); });")>]
let private onUnhandledRejection (handler: obj -> unit) : unit = jsNative

[<Emit("process.on('uncaughtException', function (error) { $0(error); });")>]
let private onUncaughtException (handler: obj -> unit) : unit = jsNative
#else
let private onUnhandledRejection (_: obj -> unit) : unit = ()
let private onUncaughtException (_: obj -> unit) : unit = ()
#endif

let all =
    testList "All"
        [
            Workbook.Tests.main
            DefaultIO.Tests.main
            Json.Tests.main
        ]


[<EntryPoint>]
let main argv =
    let unhandledMessages = ResizeArray<string>()

    onUnhandledRejection (fun reason ->
        let msg = errorToString reason
        unhandledMessages.Add($"[unhandledRejection]{System.Environment.NewLine}{msg}")
        printfn "%s" msg
    )

    onUncaughtException (fun error ->
        let msg = errorToString error
        unhandledMessages.Add($"[uncaughtException]{System.Environment.NewLine}{msg}")
        printfn "%s" msg
    )

    let exitCode = Pyxpecto.runTests [||] all
    if unhandledMessages.Count > 0 then
        printfn "Unhandled async errors captured: %i" unhandledMessages.Count
        unhandledMessages
        |> Seq.iter (fun msg -> printfn "%s" msg)
        if exitCode = 0 then 2 else exitCode
    else
        exitCode