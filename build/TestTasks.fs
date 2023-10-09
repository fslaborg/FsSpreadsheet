module TestTasks

open BlackFox.Fake
open Fake.DotNet

open ProjectInfo
open BasicTasks

[<Literal>]
let FableTestPath_input = "tests/FsSpreadsheet.Tests"

module RunTests = 

    /// runs `npm test` in root. 
    /// npm test consists of `test` and `pretest`
    /// check package.json in root for behavior
    let runTestsJs = BuildTask.create "runTestsJS" [clean; build] {
        run npm "test" ""
        run npm "run testexceljs" ""
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