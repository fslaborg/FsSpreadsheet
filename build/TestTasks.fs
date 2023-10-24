﻿module TestTasks

open BlackFox.Fake
open Fake.DotNet

open ProjectInfo
open BasicTasks

[<Literal>]
let FableTestPath_input = "tests/FsSpreadsheet.Tests"

module RunTests = 

    open Fake.Core

    //let createFreshTestFiles = BuildTask.create "createFreshTestFiles" [] {
    //    let testFilesPath = "./tests/TestUtils/TestFiles"
    //    let source = System.IO.FileInfo(testFilesPath + @"/TestWorkbook_Excel.xlsx")
    //    let scriptsFolder = "/Scripts"
    //    let testFiles = 
    //        [|
    //            @"/TestWorkbook_FsSpreadsheet.net.xlsx",    @".\runFsSpreadsheet.fsx.cmd"
    //            @"/TestWorkbook_FsSpreadsheet.js.xlsx",     @".\runFsSpreadsheet.js.cmd"
    //            @"/TestWorkbook_FableExceljs.xlsx",         @".\runFableExceljs"
    //            @"/TestWorkbook_ClosedXML.xlsx",            @".\runClosedXml"
    //        |]

    //    for testFile, script in testFiles do
    //        let target = System.IO.FileInfo(testFilesPath + testFile)
    //        if source.LastWriteTimeUtc > target.LastWriteTimeUtc then
    //            let scriptFolderPath = testFilesPath + scriptsFolder
    //            Trace.traceImportant $"Update `{testFile}` with `{script}`, as source file was updated since last transpilation."
    //            run (createProcess script) "" scriptFolderPath
    //}

    /// runs `npm test` in root. 
    /// npm test consists of `test` and `pretest`
    /// check package.json in root for behavior
    let runTestsJs = BuildTask.create "runTestsJS" [clean; build] {
        run npm "test" ""
        run npm "run testexceljs" ""
        run npm "run testjs" ""
    }

    let runTestsDotnet = BuildTask.create "runTestsDotnet" [clean; build] {
        testProjects
        |> Seq.iter (fun testProject ->
            Fake.DotNet.DotNet.test(fun testParams ->
                {
                    testParams with
                        Logger = Some "console;verbosity=detailed"
                        Configuration = DotNet.BuildConfiguration.fromString configuration
                        NoBuild = true
                }
            ) testProject
        )
    }

let runTests = BuildTask.create "RunTests" [clean; build; RunTests.runTestsJs; RunTests.runTestsDotnet] { 
    ()
}