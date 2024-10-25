module BasicTasks

open BlackFox.Fake
open Fake.IO
open Fake.DotNet
open Fake.IO.Globbing.Operators

open ProjectInfo

[<AutoOpen>]
module Helper =

    open Fake
    open Fake.Core

    let createProcess exe arg dir =
        CreateProcess.fromRawCommandLine exe arg
        |> CreateProcess.withWorkingDirectory dir
        |> CreateProcess.ensureExitCode

    module Proc =

        module Parallel =

            open System

            let locker = obj()

            let colors = [|  
                ConsoleColor.DarkYellow
                ConsoleColor.DarkCyan 
                ConsoleColor.Magenta
                ConsoleColor.Blue
                ConsoleColor.Cyan
                ConsoleColor.DarkMagenta
                ConsoleColor.DarkBlue
                ConsoleColor.Yellow
            |]

            let print color (colored: string) (line: string) =
                lock locker
                    (fun () ->
                        let currentColor = Console.ForegroundColor
                        Console.ForegroundColor <- color
                        Console.Write colored
                        Console.ForegroundColor <- currentColor
                        Console.WriteLine line)

            let onStdout index name (line: string) =
                let color = colors.[index % colors.Length]
                if isNull line then
                    print color $"{name}: --- END ---" ""
                else if String.isNotNullOrEmpty line then
                    print color $"{name}: " line

            let onStderr name (line: string) =
                let color = ConsoleColor.Red
                if isNull line |> not then
                    print color $"{name}: " line

            let redirect (index, (name, createProcess)) =
                createProcess
                |> CreateProcess.redirectOutputIfNotRedirected
                |> CreateProcess.withOutputEvents (onStdout index name) (onStderr name)

            let printStarting indexed =
                for (index, (name, c: CreateProcess<_>)) in indexed do
                    let color = colors.[index % colors.Length]
                    let wd =
                        c.WorkingDirectory
                        |> Option.defaultValue ""
                    let exe = c.Command.Executable
                    let args = c.Command.Arguments.ToStartInfo
                    print color $"{name}: {wd}> {exe} {args}" ""

            let run cs =
                cs
                |> Seq.toArray
                |> Array.indexed
                |> fun x -> printStarting x; x
                |> Array.map redirect
                |> Array.Parallel.map Proc.run

    let dotnet = createProcess "dotnet"

    let npm =
        let npmPath =
            match ProcessUtils.tryFindFileOnPath "npm" with
            | Some path -> path
            | None ->
                "npm was not found in path. Please install it and make sure it's available from your path. " +
                "See https://safe-stack.github.io/docs/quickstart/#install-pre-requisites for more info"
                |> failwith

        createProcess npmPath

    let python = 
        if System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) then
            Fake.Core.Trace.log "Detected Windows System."
            createProcess (__SOURCE_DIRECTORY__.Replace(@"\build",@"\.venv\Scripts\python.exe"))
        else
            Fake.Core.Trace.log "Detected Unix System."
            createProcess (__SOURCE_DIRECTORY__.Replace(@"/build",@"/.venv/bin/python"))

    let run proc arg dir =
        proc arg dir
        |> Proc.run
        |> ignore

    let runParallel processes =
        processes
        |> Proc.Parallel.run
        |> ignore

let setPrereleaseTag = BuildTask.create "SetPrereleaseTag" [] {
    printfn "Please enter pre-release package suffix"
    let suffix = System.Console.ReadLine()
    prereleaseSuffix <- suffix
    prereleaseTag <- (sprintf "%i.%i.%i-%s" release.SemVer.Major release.SemVer.Minor release.SemVer.Patch suffix)
    isPrerelease <- true
}

let clean = BuildTask.create "Clean" [] {
    !! "src/**/bin"
    ++ "src/**/obj"
    ++ "tests/**/bin"
    ++ "tests/**/obj"
    ++ "dist"
    |> Shell.cleanDirs 
}

let build = BuildTask.create "Build" [clean] {
        solutionFile 
        |> DotNet.build (fun p ->
            let msBuildParams =
                {p.MSBuildParams with 
                    Properties = ([
                        "warnon", "3390"
                    ])
                    DisableInternalBinLog = true
                }
            {
                p with 
                    MSBuildParams = msBuildParams
            }
            |> DotNet.Options.withCustomParams (Some "-tl")
        )
    }