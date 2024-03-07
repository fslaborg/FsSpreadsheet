module PackageTasks

open MessagePrompts
open BasicTasks
open TestTasks

open BlackFox.Fake
open Fake.Core
open Fake.IO.Globbing.Operators

open System.Text.RegularExpressions

/// https://github.com/Freymaurer/Fake.Extensions.Release#release-notes-in-nuget
let private replaceCommitLink input = 
    let commitLinkPattern = @"\[\[#[a-z0-9]*\]\(.*\)\] "
    Regex.Replace(input,commitLinkPattern,"")

module BundleDotNet =
    let bundle (versionTag : string) (versionSuffix : string option) =
        System.IO.Directory.CreateDirectory(ProjectInfo.netPkgDir) |> ignore
        !! "src/**/*.*proj"
        -- "src/bin/*"
        |> Seq.iter (Fake.DotNet.DotNet.pack (fun p ->
            let msBuildParams =
                {p.MSBuildParams with 
                    Properties = ([
                        "Version",versionTag
                        "PackageReleaseNotes",  (ProjectInfo.release.Notes |> List.map replaceCommitLink |> String.toLines )
                    ] @ p.MSBuildParams.Properties)
                }
            {
                p with 
                    VersionSuffix = versionSuffix
                    MSBuildParams = msBuildParams
                    OutputPath = Some ProjectInfo.netPkgDir
            }
        ))

let packDotNet = BuildTask.create "PackDotNet" [clean; build; runTests] {
    BundleDotNet.bundle ProjectInfo.stableVersionTag None
}

let packDotNetPrerelease = BuildTask.create "PackDotNetPrerelease" [setPrereleaseTag; clean; build; runTests] {
    BundleDotNet.bundle ProjectInfo.prereleaseTag (Some ProjectInfo.prereleaseTag)
}

module BundleJs =
    let bundle (versionTag) =
        System.IO.Directory.CreateDirectory(ProjectInfo.npmPkgDir) |> ignore
        run npm "run bundle" ""
        Fake.IO.File.readAsString "package.json"
        |> fun t ->
            let t = t.Replace(ProjectInfo.stableVersionTag, versionTag)
            Fake.IO.File.writeString false $"{ProjectInfo.npmPkgDir}/package.json" t
        "" // "fable-library.**/**"
        |> Fake.IO.File.writeString false $"{ProjectInfo.npmPkgDir}/fable_modules/.npmignore"

let packJS = BuildTask.create "PackJS" [clean; build; runTests] {
    BundleJs.bundle ProjectInfo.stableVersionTag
}

let packJSPrerelease = BuildTask.create "PackJSPrerelease" [setPrereleaseTag; clean; build; runTests] {
    BundleJs.bundle ProjectInfo.prereleaseTag
    }

module BundlePy =
    let bundle (versionTag: string) =
        
        run dotnet $"fable src/FsSpreadsheet.Py -o {ProjectInfo.pyPkgDir}/fsspreadsheet --lang python" ""
        run python "-m poetry install --no-root" ProjectInfo.pyPkgDir
        //GenerateIndexPy.ARCtrl_generate (ProjectInfo.pyPkgDir + "/arctrl")

        Fake.IO.File.readAsString "pyproject.toml"
        |> fun t ->
            let t = t.Replace(ProjectInfo.stableVersionTag, versionTag)
            Fake.IO.File.writeString false $"{ProjectInfo.pyPkgDir}/pyproject.toml" t

        Fake.IO.File.readAsString "README.md"
        |> Fake.IO.File.writeString false $"{ProjectInfo.pyPkgDir}/README.md"

        run python "-m poetry build" ProjectInfo.pyPkgDir //Remove "-o ." because not compatible with publish 
        
let packPy = BuildTask.create "PackPy" [clean; build; runTests] {
    BundlePy.bundle ProjectInfo.stableVersionTag
}

let pack = BuildTask.createEmpty "Pack" [packDotNet; packJS; packPy]

let packPrerelease = BuildTask.createEmpty "PackPrerelease" [packDotNetPrerelease; packJSPrerelease]