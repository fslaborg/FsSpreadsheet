module PackageTasks

open ProjectInfo

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

let packDotNet = BuildTask.create "PackDotNet" [clean; build; runTests] {
    if promptYesNo (sprintf "[.NET] creating stable package with version %s OK?" stableVersionTag ) 
        then
            !! "src/**/*.*proj"
            -- "src/bin/*"
            |> Seq.iter (Fake.DotNet.DotNet.pack (fun p ->
                let msBuildParams =
                    {p.MSBuildParams with 
                        Properties = ([
                            "Version",stableVersionTag
                            "PackageReleaseNotes",  (release.Notes |> List.map replaceCommitLink |> String.concat "\r\n" )
                        ] @ p.MSBuildParams.Properties)
                    }
                {
                    p with 
                        MSBuildParams = msBuildParams
                        OutputPath = Some pkgDir
                }
            ))
    else failwith "aborted"
}

let packDotNetPrerelease = BuildTask.create "PackDotNetPrerelease" [setPrereleaseTag; clean; build; runTests] {
    if promptYesNo (sprintf "[.NET] creating prerelease package with version %s OK?" prereleaseTag )
        then 
            !! "src/**/*.*proj"
            -- "src/bin/*"
            |> Seq.iter (Fake.DotNet.DotNet.pack (fun p ->
                let msBuildParams =
                    {p.MSBuildParams with 
                        Properties = ([
                            "Version", prereleaseTag
                            "PackageReleaseNotes",  (release.Notes |> List.map replaceCommitLink  |> String.toLines )
                        ] @ p.MSBuildParams.Properties)
                    }
                {
                    p with 
                        VersionSuffix = Some prereleaseSuffix
                        OutputPath = Some pkgDir
                        MSBuildParams = msBuildParams
                }
            ))
    else
        failwith "aborted"
}

module BundleJs =
    let bundle (versionTag) =
        run npm "run bundle" ""
        Fake.IO.File.readAsString "package.json"
        |> fun t ->
            let t = t.Replace(ProjectInfo.stableVersionTag, versionTag)
            Fake.IO.File.writeString false $"{ProjectInfo.npmPkgDir}/package.json" t
        "" // "fable-library.**/**"
        |> Fake.IO.File.writeString false $"{ProjectInfo.npmPkgDir}/fable_modules/.npmignore"

let packJS = BuildTask.create "PackJS" [clean; build; runTests] {
    if promptYesNo (sprintf "[NPM] creating stable package with version %s OK?" ProjectInfo.stableVersionTag ) 
        then
            BundleJs.bundle ProjectInfo.stableVersionTag
    else failwith "aborted"
}

let packJSPrerelease = BuildTask.create "PackJSPrerelease" [setPrereleaseTag; clean; build; runTests] {
    if promptYesNo (sprintf "[NPM] creating prerelease package with version %s OK?" ProjectInfo.prereleaseTag ) then
        BundleJs.bundle ProjectInfo.prereleaseTag
    else failwith "aborted"
    }
let pack = BuildTask.createEmpty "Pack" [packDotNet; packJS]

let packPrerelease = BuildTask.createEmpty "PackPrerelease" [packDotNetPrerelease;packJSPrerelease]