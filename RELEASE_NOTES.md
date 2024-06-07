### 6.1.3+3ba93ef (Released 2024-6-7)
* Additions:
    * [[#9c30d63](https://github.com/CSBiology/FsSpreadsheet/commit/9c30d637eabfb6eca070a105bd39bf9d177d423b)] add manage-issues workflow :green_heart:
    * [[#1ab4da0](https://github.com/CSBiology/FsSpreadsheet/commit/1ab4da0a7abf1b724230a75f108a6d71a06a739f)] Merge pull request #89 from fslaborg/pythonWorkbookReading
    * [[#7fe1c6a](https://github.com/CSBiology/FsSpreadsheet/commit/7fe1c6ac6fd2d31c3278dfee498828d8c03a9833)] bump to 6.1.2
* Bugfixes:
    * [[#3ba93ef](https://github.com/CSBiology/FsSpreadsheet/commit/3ba93efadde0aa73184d52bf38405f82d97ecfa5)] Fix missing fable includes #90

### 6.1.2+62db5ab (Released 2024-4-9)
* Additions:
    * [[#62db5ab](https://github.com/CSBiology/FsSpreadsheet/commit/62db5abb62cfd411355f5e4b6a466aeb622e448e)] update project dependenices on FsSpreadsheet.Core

### 6.1.1+0f0535a (Released 2024-4-8)
* Additions:
    * [[#0d578e2](https://github.com/CSBiology/FsSpreadsheet/commit/0d578e27d65179857267c0f4601bc75a22327b9a)] bump to 6.1.0
* Bugfixes:
    * [[#0f0535a](https://github.com/CSBiology/FsSpreadsheet/commit/0f0535a4afbc44e5980c89e608334b96e38a4a7d)] check and fix py xlsx reader rescanning rows

### 6.1.0+ed6795c (Released 2024-3-19)
* Additions:
    * [[#3f1fb55](https://github.com/CSBiology/FsSpreadsheet/commit/3f1fb55e75d2ea56dec8787f0e6dfceccfd19ef5)] start working on json writing
    * [[#13dbca2](https://github.com/CSBiology/FsSpreadsheet/commit/13dbca29d9082d91a47df04c468c89fe1315e405)] finish first version of json io functions
    * [[#514c204](https://github.com/CSBiology/FsSpreadsheet/commit/514c20456a6b49c2c66d35533a728d60e7d2faa2)] add json io tests
    * [[#02baf70](https://github.com/CSBiology/FsSpreadsheet/commit/02baf70c27534b670d1c1d41d744c8ecb0adecec)] naming adjustments and docs update
    * [[#a6c882c](https://github.com/CSBiology/FsSpreadsheet/commit/a6c882ca9ee2b0155aa9894599900f52488e48ac)] set ziparchivereader performance test to pending
    * [[#9ceda1e](https://github.com/CSBiology/FsSpreadsheet/commit/9ceda1e6ba9705b13d4a9f65f90eb89eeeb8cc27)] add old xlsx IO functions with deprecation attribute
    * [[#5e71fb9](https://github.com/CSBiology/FsSpreadsheet/commit/5e71fb9bbd549e2a06371d10e9213a740ec096c4)] add try json parse functions
* Bugfixes:
    * [[#1586c5a](https://github.com/CSBiology/FsSpreadsheet/commit/1586c5a04bd35dfb0532801bcde2b1e2ca714a8b)] fix json io datetime issues

### 6.0.1+360ee6f (Released 2024-3-7)
* Additions:
    * Hotfix because of erroneous namespace names
    * [[#c6aa628](https://github.com/CSBiology/FsSpreadsheet/commit/c6aa6281c0adbc7436227f99db69b19aa25d8707)] bump to 6.0.0
    * [[#c7b4972](https://github.com/CSBiology/FsSpreadsheet/commit/c7b4972ce5921c259d18bbc7a4ee8adc7139cdca)] updated namespace names according to project name unification
    * [[#360ee6f](https://github.com/CSBiology/FsSpreadsheet/commit/360ee6f12867663c4984d8d730389a830d99bc2e)] add minimal documentation

### 6.0.0+b85ca85 (Released 2024-3-7)
* Additions:
    * Added release for python
    * Renamed xlsx io packages
    * [[#9572480](https://github.com/CSBiology/FsSpreadsheet/commit/9572480de9ae43591714e0cff6ee92a00b0541b0)] rename excelpy folder
    * [[#491ed1e](https://github.com/CSBiology/FsSpreadsheet/commit/491ed1ee1bc3cafdfc4538e98dd965aaf8a3dad0)] update CI and add Python build targets
    * [[#9525665](https://github.com/CSBiology/FsSpreadsheet/commit/95256651d8930ffd7645edeacf2cd731a64ab276)] update readme with developer info
    * [[#181fc70](https://github.com/CSBiology/FsSpreadsheet/commit/181fc703db651864d84295648a1db9fa3bae1e19)] unify test project names
    * [[#b85ca85](https://github.com/CSBiology/FsSpreadsheet/commit/b85ca85ca290076d21ae6af01ab3aab0cd52205a)] small changes to reduce warnings

### 5.2.0+c6f526ad (Released 2024-2-23)
* Additions:
    * [[#bbba6d4](https://github.com/CSBiology/FsSpreadsheet/commit/bbba6d4ab819900f5188af3b5663d4744061991a)] start setup of python package
    * [[#80dd1b6](https://github.com/CSBiology/FsSpreadsheet/commit/80dd1b608e46e7cda7646ea7c39619c0c7738f8e)] add tests for python xlsx io
    * [[#992fa5a](https://github.com/CSBiology/FsSpreadsheet/commit/992fa5ac485313efc71285347b156b83614585d0)] add python xlsx io support
    * [[#ddcc16d](https://github.com/CSBiology/FsSpreadsheet/commit/ddcc16dbdc51da516f53cbc3192d114584bd73fb)] add runtestspy target
    * [[#d69288a](https://github.com/CSBiology/FsSpreadsheet/commit/d69288a33f1a259987cf4c229671723eaaf5beb2)] switch to pyxpecto for testing
    * [[#2b88258](https://github.com/CSBiology/FsSpreadsheet/commit/2b882586b12aae0500eaad98d01603521dfee26f)] small speed improvement of reader by skipping sst include on opendocument cell
* Bugfixes:
    * [[#b35f309](https://github.com/CSBiology/FsSpreadsheet/commit/b35f30908a4a68d8e3870c30cb8ace2dde748a7b)] finalize python xlsx io test and fixes
    * [[#a55336e](https://github.com/CSBiology/FsSpreadsheet/commit/a55336ec03d5d2f1a254d2908d06d9b5381c7d55)] fix python xlsx io for default formats

### 5.1.3+a260549 (Released 2024-2-13)
* Additions:
    * [[#a260549](https://github.com/CSBiology/FsSpreadsheet/commit/a26054928881975896f2c07926261091979216a4)] update tests to change in empty workbook handling
    * [[#4eaf806](https://github.com/CSBiology/FsSpreadsheet/commit/4eaf806c4e64bcabc762517e07001886efde4611)] add fail for writing empty workbook #38
    * [[#6cc4d8d](https://github.com/CSBiology/FsSpreadsheet/commit/6cc4d8dfaa91832d6e28fdffe638600570c4e6fd)] Merge pull request #84 from fslaborg/speed
    * [[#c9d00fb](https://github.com/CSBiology/FsSpreadsheet/commit/c9d00fb5203377447c25ef475d85e74c428cb6a9)] some speed improvements to xlsx reading
    * [[#bd9b64b](https://github.com/CSBiology/FsSpreadsheet/commit/bd9b64be581a772513391b1758cfa1e78ab01473)] first draft of experimental ziparchive xlsx reader
    * [[#2b88258](https://github.com/CSBiology/FsSpreadsheet/commit/2b882586b12aae0500eaad98d01603521dfee26f)] small speed improvement of reader by skipping sst include on opendocument cell
* Bugfixes:
    * [[#8da67c9](https://github.com/CSBiology/FsSpreadsheet/commit/8da67c9c0ac537ba184d22b16f86d2d1b18f17ab)] small fixes and adjustments to xlsx reader and tests
    * [[#886dcf2](https://github.com/CSBiology/FsSpreadsheet/commit/886dcf2ef7459096bf1a35a4fc6173e33f424758)] fix zipArchiveReader against default tests

### 5.1.2+21fd6e8 (Released 2024-1-30)
* Additions:
    * [[#21fd6e8](https://github.com/CSBiology/FsSpreadsheet/commit/21fd6e82f6d9eab48b4197a56a5c9b9f06b51b0e)] improve reader speed by adjusting sharedStringTable usage
* Bugfixes:
    * [[#66b6713](https://github.com/CSBiology/FsSpreadsheet/commit/66b6713381be3f133c58db7ca78fd36bd63c0ad1)] fix npm publish target

### 5.1.1+e7cc638 (Released 2024-1-26)
* Additions:
    * [[#e7cc638](https://github.com/CSBiology/FsSpreadsheet/commit/e7cc638b170c570005b6cd8378d1e0ac31075be7)] improve rowWithRange SkipSearch performance

### 5.1.0+4732d22 (Released 2024-1-24)
* Additions:
    * Increase Parser speed
    * [[#4732d22](https://github.com/CSBiology/FsSpreadsheet/commit/4732d22e6acbf70b284355149fd4d06932023783)] include parse speed test in js
    * [[#9aa368f](https://github.com/CSBiology/FsSpreadsheet/commit/9aa368f3cefeaa50c9eb4c13865bb952a57a273c)] add speedtest for excel file reader
    * [[#f401c96](https://github.com/CSBiology/FsSpreadsheet/commit/f401c9618bfd66648688cfbb8b99e51f46b14bdd)] make RescanRows speed linear

### 5.0.2+41eca2f (Released 2023-11-3)
* Bugfixes:
    * [[#b45db41](https://github.com/CSBiology/FsSpreadsheet/commit/b45db41a89e78576999afd28a2d7b1e087caf335)] Fxi critical bug in fable compatibility

### 5.0.1+2958cd9 (Released 2023-10-28)
* Additions:
    * [[#2958cd9](https://github.com/CSBiology/FsSpreadsheet/commit/2958cd951db833c5d469875d03ad175e496f07ae)] update references from major 3 to major 5
    * [[#dff1bf3](https://github.com/CSBiology/FsSpreadsheet/commit/dff1bf35cafa238fbe3604524259684aa90f3520)] Update version and release npm

### 5.0.0+62ff8ec (Released 2023-10-26)
* Additions:
    * [[#db263ed](https://github.com/CSBiology/FsSpreadsheet/commit/db263edea5bd8ecd9e6e8023fdae6a396fa01cec)] Start implementing logic for auto update, but push to backlog for now
    * [[#c5bf348](https://github.com/CSBiology/FsSpreadsheet/commit/c5bf34808d105723738dde53e6cb778a2b97da3e)] Finish io tests :heavy_check_mark:
    * [[#80f7632](https://github.com/CSBiology/FsSpreadsheet/commit/80f7632b2a4865951f144fc271e02a2ee4f6989c)] Improve DateTime numberFormat recognition
    * [[#baa88ce](https://github.com/CSBiology/FsSpreadsheet/commit/baa88ceb185c43302c8e7742ba8fe1afab893739)] update libre test file
    * [[#9805811](https://github.com/CSBiology/FsSpreadsheet/commit/980581193cdcceaa80cd87f91e86121726c1046f)] Setup current FsSpreadsheet.Js write for defaultio tests :construction:
    * [[#05e0ca4](https://github.com/CSBiology/FsSpreadsheet/commit/05e0ca492e3ae54cdaf6fdec4c19c1296614e565)] update gitignore
    * [[#2b99506](https://github.com/CSBiology/FsSpreadsheet/commit/2b99506e9abafd7d579d8e3f09fb59d7e20a8398)] Further work on #71 change FsCell.Value to obj :boom:
    * [[#546e029](https://github.com/CSBiology/FsSpreadsheet/commit/546e02962c5d18127e143d85705523a4163d63c2)] add stylesheet helper functions
    * [[#47615d7](https://github.com/CSBiology/FsSpreadsheet/commit/47615d7cfba7b19eb53815ac500c2dfd45a02d3a)] add preliminary stylesheet functionality for cell datetime writing (wip)
    * [[#1a71e96](https://github.com/CSBiology/FsSpreadsheet/commit/1a71e9688f02d5d35e89f3cacf380f99d5049195)] Update files correctly
    * [[#326cbba](https://github.com/CSBiology/FsSpreadsheet/commit/326cbbafda241edc6ecd85b46247969b100cc56d)] Make first tests pass :heavy_check_mark:
    * [[#085418f](https://github.com/CSBiology/FsSpreadsheet/commit/085418fa77844685613e55e5542f91d95d7ace85)] Add FsSpreadsheet tests and update testFile to include Time in DateTime
    * [[#4d7c27f](https://github.com/CSBiology/FsSpreadsheet/commit/4d7c27f78917a445cd826fdd6bde03211d8ea636)] Add boolean column to Default TestWorkbook
    * [[#371808c](https://github.com/CSBiology/FsSpreadsheet/commit/371808cecb3f5255fb58dae4b2f3bf317090753e)] Make ExcelIO FableExceljs and OpenXML tests pass :heavy_check_mark:
    * [[#759c7a0](https://github.com/CSBiology/FsSpreadsheet/commit/759c7a0ccfbe7f742cb6d3fdf87717e084c29329)] Update files from TestWorkbook_Excel
    * [[#58736f3](https://github.com/CSBiology/FsSpreadsheet/commit/58736f384a1fa98254ede83e643c67ce2d239cbd)] make ExcelIO-Excel test pass :heavy_check_mark:
    * [[#f468fd0](https://github.com/CSBiology/FsSpreadsheet/commit/f468fd04e5de0761ea376e6e4ca9708aa321b194)] Parse correct DataType from OpenXML #73
    * [[#1ad6c9d](https://github.com/CSBiology/FsSpreadsheet/commit/1ad6c9d960701b4f565a0abd210ea8fd3845e49f)] Move TestFiles and setup Utility project #71
    * [[#4cc9c25](https://github.com/CSBiology/FsSpreadsheet/commit/4cc9c25522be978283d2bacbd16dd19bcdf6f509)] Improve HeaderRow functions #72
    * [[#b8a51bf](https://github.com/CSBiology/FsSpreadsheet/commit/b8a51bf7eedf08e48460a302c25954e4807daa27)] add testFile info
    * [[#88cffd2](https://github.com/CSBiology/FsSpreadsheet/commit/88cffd2ad2e6d05e213c9426f1a812e78b10843d)] Add io helper scripts
    * [[#0811a1d](https://github.com/CSBiology/FsSpreadsheet/commit/0811a1d0b211c34cb8bbd65cf13f3df5f574b706)] update xlsx file
    * [[#cfdb888](https://github.com/CSBiology/FsSpreadsheet/commit/cfdb888b02e2a3d357e6bf70434ba97180532016)] add excel testTable
    * [[#eedf32f](https://github.com/CSBiology/FsSpreadsheet/commit/eedf32f1e05d2ccba3b259bc92d908187b48ae25)] Merge pull request #70 from fslaborg/issue_69
    * [[#858d6d8](https://github.com/CSBiology/FsSpreadsheet/commit/858d6d83f08c3d70db017c3c0964e4fd6aae01a3)] add correct extensions to native js tests
    * [[#a54d874](https://github.com/CSBiology/FsSpreadsheet/commit/a54d874f1345117dcf0e335e522b74780619053f)] Set correct default for showHeaderRow = false #69
    * [[#e0b42c7](https://github.com/CSBiology/FsSpreadsheet/commit/e0b42c7ac4394c19105997b8892c2966e5f584ee)] Merge pull request #68 from fslaborg/ioExtension
    * [[#e903e76](https://github.com/CSBiology/FsSpreadsheet/commit/e903e767c1d1e22f000778b849ade4c4587cfe68)] add additional xlsx readers variants
    * [[#a4cd326](https://github.com/CSBiology/FsSpreadsheet/commit/a4cd326f7166dcd7b59a56f2f83cadf020339114)] Merge branch 'main' of https://github.com/CSBiology/FsSpreadsheet
    * [[#c69f787](https://github.com/CSBiology/FsSpreadsheet/commit/c69f7877f9cb8c96ad995a879e6d3a310407635a)] Merge pull request #66 from fslaborg/npmRelease
    * [[#63498a9](https://github.com/CSBiology/FsSpreadsheet/commit/63498a9350d9f46c356c6bf73c98c1b30ec1fb95)] Add tasks for npm release
    * [[#9c7ab69](https://github.com/CSBiology/FsSpreadsheet/commit/9c7ab691b4aca6c0b16d0c2c7fd374b8d36e14e9)] Add test to verify correct parsing of old ClosedXml
    * [[#02d92c2](https://github.com/CSBiology/FsSpreadsheet/commit/02d92c299e2e1358a201f1617d2853e0f885cb86)] Update exceljs dependency
    * [[#9eb94ea](https://github.com/CSBiology/FsSpreadsheet/commit/9eb94eabf3c91d1bb7390e40c11e41dc65906893)] Update package.json: Mocha timeout for Exceljs JS tests
    * [[#f105a75](https://github.com/CSBiology/FsSpreadsheet/commit/f105a75079f14bc3408984d46b5bc69d420a038d)] Merge pull request #64 from CSBiology/feature-moreFsColumnFun-#63
    * [[#36a72a2](https://github.com/CSBiology/FsSpreadsheet/commit/36a72a26b5f865031a31e9ef1fc47d6216ed1935)] Update package.json, update and rewrite Release Notes for v4.0.0
    * [[#d501225](https://github.com/CSBiology/FsSpreadsheet/commit/d50122572e9310d1fbab0b971e8922a884775f38)] Add unit tests for newly added functions
* Bugfixes:
    * [[#56d533f](https://github.com/CSBiology/FsSpreadsheet/commit/56d533fd3279f8d6c7e6f56a489e1b734abcb166)] Fix ExcelIO: bool & float write. Fix exceljs: table postion write, cell value type read and write, confusing naming. Add scripts for fsspreadsheet write. :sparkles:🤡
    * [[#6fe7f89](https://github.com/CSBiology/FsSpreadsheet/commit/6fe7f891644934b427ea56620fa8983befc25b10)] Fix ci
    * [[#17a9cf8](https://github.com/CSBiology/FsSpreadsheet/commit/17a9cf8e51a8b1f3fa7e11461f923bddff1ca220)] Fix embedded ressource path
    * [[#26c574d](https://github.com/CSBiology/FsSpreadsheet/commit/26c574dd29e0908d00cdb8e33c87d8e142052361)] fix libre file parsing and update test files
    * [[#47e696b](https://github.com/CSBiology/FsSpreadsheet/commit/47e696ba6e25f1b02e3d520bc0c77e7df834a3df)] update spreadsheet file and final fixes for dotnet reader
    * [[#5f6c8da](https://github.com/CSBiology/FsSpreadsheet/commit/5f6c8da5c438458d8838fab48612b691db6ddd23)] Fix DateTime format :bug:
    * [[#d6e85cb](https://github.com/CSBiology/FsSpreadsheet/commit/d6e85cba08422c17bf323577b4ccea729bbe97a7)] Fix DateTime and paths for build chain
    * [[#3a60e29](https://github.com/CSBiology/FsSpreadsheet/commit/3a60e29268d2f65646d751088d8ca09f841125aa)] Fix paths, but exceljs parses DateTime differently :construction:
    * [[#77b722d](https://github.com/CSBiology/FsSpreadsheet/commit/77b722d2dd6c15836ef23645e9c3958861d69309)] Fix read in for boolean type `DataType` :bug:
    * [[#c7b88b6](https://github.com/CSBiology/FsSpreadsheet/commit/c7b88b6a909700da7fa2c481462435930fc2e62f)] try fix CI
    * [[#116962e](https://github.com/CSBiology/FsSpreadsheet/commit/116962e42a2b8e52128d4d039dc50d86dd4588ab)] Fix error in fsspreadsheet not writing table.name only displayname
    * [[#2ea7401](https://github.com/CSBiology/FsSpreadsheet/commit/2ea7401b78d91a4f75a9cb77a033142610de21d5)] fix package.json publish bug :bug:

### 4.0.0+5d5067f (Released 2023-8-11)
* Additions:
    * [[#19bd2a8](https://github.com/CSBiology/FsSpreadsheet/commit/19bd2a8817940d4f8725a85ebfa595faa214a24d)] Add `MinRowIndex` and `MinColIndex` to FsColumn
    * [[#5c2e616](https://github.com/CSBiology/FsSpreadsheet/commit/5c2e6167c3c0ea9bfd6ae4a59de60485325a31ab)] Add `HasCellAt` to FsColumn
    * [[#3e7b7d2](https://github.com/CSBiology/FsSpreadsheet/commit/3e7b7d2c76a7ab1b72c13dd20983b975cb8b1308)] Add `TryItem` functionality to FsColumn
    * [[#5d5067f](https://github.com/CSBiology/FsSpreadsheet/commit/5d5067f81459164c72e6abf7db1c94b7b494d1ee)] Add dense column functionality
    * [[#ac14e0d](https://github.com/CSBiology/FsSpreadsheet/commit/ac14e0d2b8df599aa83e0c93326553f4bb2c2d79)] Add unit tests for column index getting and base cells in FsRow
    * [[#36b2b16](https://github.com/CSBiology/FsSpreadsheet/commit/36b2b163e17a944da64f04074e539cc7cd687b0c)] Add tests for FsRow
    * [[#05f19ae](https://github.com/CSBiology/FsSpreadsheet/commit/05f19ae60db21031fb0ee527ef94431d199bee6e)] Add `GetRowCount` functionality to FsRangeBase
    * [[#aa4938c](https://github.com/CSBiology/FsSpreadsheet/commit/aa4938cb7dff34e84cf4d15d7325fd60ed1ebd85)] Add `GetRows` functionality to FsTable
    * [[#2a29cfe](https://github.com/CSBiology/FsSpreadsheet/commit/2a29cfe16f4360ad0f0fbb213ccc0abcbb653087)] Add FsRow members: `TryItem`, `ToDenseRow`, `HasCellAt`, `MinColIndex`, `MaxColIndex`
    * 	* [[#c228e7f](https://github.com/CSBiology/FsSpreadsheet/commit/c228e7f3799d87d34e44fb96500fe4f360d6454e)] Change npm dependency to exceljs fork
    * [[#5625b33](https://github.com/CSBiology/FsSpreadsheet/commit/5625b3310a5918a5396d49a7bef1e6654d50d295)] Update more tests
    * [[#4378808](https://github.com/CSBiology/FsSpreadsheet/commit/4378808c13d3250b90ac2311fd127b999498b2c4)] Update gitignore
    * [[#cb35186](https://github.com/CSBiology/FsSpreadsheet/commit/cb3518609e9fe42e05789441906dc148a4b8d53e)] Update and clean up RELEASE_NOTES.md
    * [[#489f6b7](https://github.com/CSBiology/FsSpreadsheet/commit/489f6b77789cca13840382d5c2dd03dc973b65df)] Add logo to Exceljs project file
    * [[#caf21db](https://github.com/CSBiology/FsSpreadsheet/commit/caf21db25490908a61381fd229fdaf5cc9a66ab2)] INcrease package.json version together with fake task.
    * [[#2b9c7fa](https://github.com/CSBiology/FsSpreadsheet/commit/2b9c7fa2c3fccfb6a5956b2e244a557fda4e0793)] Add some tests for exceljs
    * [[#1fd0e34](https://github.com/CSBiology/FsSpreadsheet/commit/1fd0e34fafc9e82f0a509af9a4d1579a4ea458d5)] Add fable necessities to fsproj.
    * [[#72bb41e](https://github.com/CSBiology/FsSpreadsheet/commit/72bb41ea6b8ad2c5490619af35a4d9038f0700da)] Write some tests #54
    * [[#95c5d16](https://github.com/CSBiology/FsSpreadsheet/commit/95c5d167403cb5a3323782b2e6895c66c9f1ebd8)] Push WIP state
    * [[#58b57ae](https://github.com/CSBiology/FsSpreadsheet/commit/58b57ae74d7d5d98e6146c75ce31eb819c0ed5d9)] Start refactoring to resizeArray :hammer:#53
    * [[#5ed3752](https://github.com/CSBiology/FsSpreadsheet/commit/5ed37521a61683dfb6f6240a1c807474f50b9574)] Specify code snippet
    * [[#4a9ff2e](https://github.com/CSBiology/FsSpreadsheet/commit/4a9ff2e91cf16c3d7d5587937f9c97d0c348e967)] Initiate test setup for FsSpreadsheet.Js
    * [[#d785795](https://github.com/CSBiology/FsSpreadsheet/commit/d785795061967c0206855983eebf49f1f182fe86)] Finish Api for both f# fable access and js native :sparkles:
    * [[#1456337](https://github.com/CSBiology/FsSpreadsheet/commit/1456337f88c8af863fb24bf134a93580ef5ed8b6)] update gitignore
    * [[#97f712c](https://github.com/CSBiology/FsSpreadsheet/commit/97f712cfe1d8422835a50c0186e28fc5cd2fcc4a)] setup msbuild
    * [[#efa9815](https://github.com/CSBiology/FsSpreadsheet/commit/efa981573a24015d110a70c2e306e2678a5a36d5)] Init FsSpreadsheet.Js and add Fable.Exceljs dependency
    * [[#c4fdce6](https://github.com/CSBiology/FsSpreadsheet/commit/c4fdce69ea59a6a7da12fd190249a8e2f8d43136)] improve testsuit
    * [[#33eb847](https://github.com/CSBiology/FsSpreadsheet/commit/33eb84701d24d145cef5801966bf25844a8a7931)] cleanup test tasks
    * [[#3116aae](https://github.com/CSBiology/FsSpreadsheet/commit/3116aae6831fddcccc552d9d68b1bf4a565578ec)] Add dotnet tool femto
    * [[#eb7160d](https://github.com/CSBiology/FsSpreadsheet/commit/eb7160d105d3d3fbabaea36abdc085412e173d75)] Update .gitignore
* Deletions:
    * [[#7779011](https://github.com/CSBiology/FsSpreadsheet/commit/77790114c7d652e818ab1c9492f730ce573c47b3)] delete dist
    * [[#2253fb1](https://github.com/CSBiology/FsSpreadsheet/commit/2253fb1ef938754f8135d18949252d991ec1af32)] Remove some js tests until fsspreadsheet syntax better fits js.
    * [[#146d920](https://github.com/CSBiology/FsSpreadsheet/commit/146d9205a549e6ccbf0b317a52fe56ea63e08bfc)] Delete .zip :fire:
    * [[#6b3a27b](https://github.com/CSBiology/FsSpreadsheet/commit/6b3a27bf8db2ee6b23c2dacce8466ea22bcda9b4)] remove annoying warning
* Bugfixes:
    * [[#2b3b30b](https://github.com/CSBiology/FsSpreadsheet/commit/2b3b30b30a03ef821668ab6a011368662622b564)] Fix critical JS bug in column index getting
    * [[#d14448e](https://github.com/CSBiology/FsSpreadsheet/commit/d14448e093ffe8268a2445d31f2aeeccf6a79fc1)] Fix Async/Promises and fsspreadsheet readin :tada:
    * [[#a96bcd0](https://github.com/CSBiology/FsSpreadsheet/commit/a96bcd0082dab203a7fdbed7ba12e96e21203f33)] fix timeout issue :bug:
    * [[#a680d45](https://github.com/CSBiology/FsSpreadsheet/commit/a680d45f9a7b985d61fa9887fd984e32024faf9a)] fix xlsxparser picking the wrong name fix teststrings
    * [[#5293ae9](https://github.com/CSBiology/FsSpreadsheet/commit/5293ae92184f4dc4038076761cd65beb3e243aba)] hotfix table not written correctly #49
    * [[#53b109c](https://github.com/CSBiology/FsSpreadsheet/commit/53b109cf42b9fba59c55011d2eb45fa39132b9c2)] fix datatypes in fs->js conversion + finish tests

### 3.1.1+64f86ec (Released 2023-7-25)
* Bugfixes:
    * [[#1ee068c](https://github.com/CSBiology/FsSpreadsheet/commit/1ee068c9c0f8f0c42137defb2837c8c54308dbc6)] fix and test #42

### 3.1.0+0d7e699 (Released 2023-7-19)
* Additions:
    * [[#af1920a](https://github.com/CSBiology/FsSpreadsheet/commit/af1920a3a05efcb0953408f9f42bb398dbb3a38e)] add attachmembers attribute to FsCell
* Bugfixes:
    * [[#d6f86d2](https://github.com/CSBiology/FsSpreadsheet/commit/d6f86d24f62c74fd0ca8aa40a6d239818256aace)] test and fix writeAndReadSpreadsheet

### 3.0.0+4da5e50 (Released 2023-7-18)
* Additions:
    * [[#a9c1f29](https://github.com/CSBiology/FsSpreadsheet/commit/a9c1f2923a12334a37b9d094025e9619fb4b4d3c)] Update test project dependencies for Fable compatability
    * [[#1c9f1bd](https://github.com/CSBiology/FsSpreadsheet/commit/1c9f1bd875af5308ff44295b284329c1ef7c5311)] use attachmembers for improved javascript transpilation
    * [[#6085c47](https://github.com/CSBiology/FsSpreadsheet/commit/6085c47c060d9ef252b9ec2fad927208f0d2bc9f)] improve comments
    * [[#4da5e50](https://github.com/CSBiology/FsSpreadsheet/commit/4da5e5062dcd1bffcb5e917585a00034dbf8bd82)] Merge pull request #34 from CSBiology/fable
* Bugfixes:
    * [[#8c8e720](https://github.com/CSBiology/FsSpreadsheet/commit/8c8e720f04af1e004cfef8c4da3236e00e7ffeb4)] Affix Build QuickStart to Readme
    * [[#70df6e6](https://github.com/CSBiology/FsSpreadsheet/commit/70df6e63578ab7330939486a81ab221af22c7eeb)] fix xml comments

### 2.0.3+3a1d082 (Released 2023-7-18)
* Additions:
    * [[#3a1d082](https://github.com/CSBiology/FsSpreadsheet/commit/3a1d08253839fe8f85f017a7a304ce766e06f71b)] Bump project version dependencies

### 2.0.2+1b8d874 (Released 2023-7-17)
* Additions:
    * latest commit #1b8d874
* Bugfixes:
    * [[#a72d301](https://github.com/CSBiology/FsSpreadsheet/commit/a72d3014417b940aeca6089bd8c973e521e2f51c)] Fix bug with read Xlsx file having no FsRows
    * [[#0dfd1f0](https://github.com/CSBiology/FsSpreadsheet/commit/0dfd1f0d11d53008ee0fe8182675a183108d7f96)] Add unit test for bug case

### 2.0.1+bd195cc (Released 2023-5-11)
* Additions:
    * [[#e8e7a88](https://github.com/CSBiology/FsSpreadsheet/commit/e8e7a88064dd85ec4d2ed6037237626fd46b7ec6)] update CI to include Fable
* Deletions:
    * [[#bd195cc](https://github.com/CSBiology/FsSpreadsheet/commit/bd195cc0920f9e723c4aa163b6935a229c7d730d)] remove deprecated functions and fill incomplete matches

### 2.0.0+2b67a3f (Released 2023-5-11)
* Additions:
    * [[#2b67a3f](https://github.com/CSBiology/FsSpreadsheet/commit/2b67a3f18a1c0cbcce92747e64e39ef4959abba7)] start adding DSL test
    * [[#9fa9537](https://github.com/CSBiology/FsSpreadsheet/commit/9fa9537ba2e9b049679ea491bc52a1e6c8707bb3)] Update ReleaseNotes
    * [[#ff35f3d](https://github.com/CSBiology/FsSpreadsheet/commit/ff35f3d060f311e6bb54ad70f80a890a49f60a49)] Add Fable logic
    * [[#d89c88b](https://github.com/CSBiology/FsSpreadsheet/commit/d89c88b126125dafc552328ec04f1c1e77036f0f)] Init fable compatibility :sparkles::tada:
    * [[#5f60c82](https://github.com/CSBiology/FsSpreadsheet/commit/5f60c8264b9655ac8c7ebb6c766458cb6506d9ec)] Add dsl logic back to fable version :sparkles:
    * [[#9b51c85](https://github.com/CSBiology/FsSpreadsheet/commit/9b51c85de78c85f2181d94e8b5be8edaa930198d)] Init fable test suit :sparkles:
    * [[#58dc08c](https://github.com/CSBiology/FsSpreadsheet/commit/58dc08c0cbbd412305e15bb1ea412e57ffae6e45)] Make tests pass :heavy_check_mark:
* Deletions:
    * [[#d158414](https://github.com/CSBiology/FsSpreadsheet/commit/d15841424793c84cd94311da7bfb59afe65c9b82)] remove argument exception mapping
* Bugfixes:
    * [[#1c72afb](https://github.com/CSBiology/FsSpreadsheet/commit/1c72afbd410bf6093ad3cc664305228fd057908e)] Fix cell getValueAs<_> for fable :sparkles:

### 1.3.0-preview+549125f (Released 2023-4-24)
* Additions:
    * [[#549125f](https://github.com/CSBiology/FsSpreadsheet/commit/549125f54d5aa1b7b41f8e52962bf506bc5eb9f4)] add Columns field to FsTable
    * [[#a43daef](https://github.com/CSBiology/FsSpreadsheet/commit/a43daefdf58e528a5a85c1542726cfa2aba4cae7)] rename members and set old counterparts as obsolete
    * [[#d25b0ac](https://github.com/CSBiology/FsSpreadsheet/commit/d25b0ac7724b8807a9f9fd62849a32cc32d07ae3)] add tests and call them with runtests target
    * [[#7840c1d](https://github.com/CSBiology/FsSpreadsheet/commit/7840c1dc2fece38de24358233344509242fbdbf8)] add IEnumerable interface to rows
    * [[#768858b](https://github.com/CSBiology/FsSpreadsheet/commit/768858bb8e55d41ef476ceb7781c3fb7bd5bd95e)] add FsColumn type
    * [[#bbed3ca](https://github.com/CSBiology/FsSpreadsheet/commit/bbed3ca4c14bbf999afd77b4cab2cf6aa4a2cd96)] Add Interactive extension with tests for formatters
    * [[#97ec087](https://github.com/CSBiology/FsSpreadsheet/commit/97ec087527b9cd21b0055b56a08fca4a9a3c03b5)] Add nuget package metadata
    * [[#571cab1](https://github.com/CSBiology/FsSpreadsheet/commit/571cab12945fffb837bf400460952089a4004dc1)] Comment
    * [[#a3de3cb](https://github.com/CSBiology/FsSpreadsheet/commit/a3de3cbae23c00f4561a64f3a86ec69efe939a72)] Update ///
    * [[#d5f599b](https://github.com/CSBiology/FsSpreadsheet/commit/d5f599bc73d14ddba8f44a1c20f987677b962b7a)] Update ///-comments & test CellValues
    * [[#38864cb](https://github.com/CSBiology/FsSpreadsheet/commit/38864cbc8322bee08493a5bceafeeecc57305c16)] Add critical error when creating FsRows
    * [[#0f443f1](https://github.com/CSBiology/FsSpreadsheet/commit/0f443f1a700d272f1942c57486151e7f68da431d)] Replace unallowed constructor pattern
* Bugfixes:
    * [[#49b5d0f](https://github.com/CSBiology/FsSpreadsheet/commit/49b5d0f5f49ddd39ddbdf6c477bcda57a341b19b)] small fix cell add function
    * [[#8ddfd39](https://github.com/CSBiology/FsSpreadsheet/commit/8ddfd39f637a2f7056f36352e08148dda7ea4a3a)] Fix typo
    * [[#808dea1](https://github.com/CSBiology/FsSpreadsheet/commit/808dea13f05c05f1a1e68d278013119f1a9188e5)] fix border
    * [[#b52ec35](https://github.com/CSBiology/FsSpreadsheet/commit/b52ec35163f247b7858d1b346b86e2d8f24ca157)] Fix critical error in FsRow constructor

### 1.0.1+99dfa53 (Released 2023-3-23)
* Additions:
    * [[#3031151](https://github.com/CSBiology/FsSpreadsheet/commit/30311515b1b0ccd6ed3828465739008b5d95fd07)] Add `tryGetById` function
    * [[#6be69e2](https://github.com/CSBiology/FsSpreadsheet/commit/6be69e22e1104731dc92fda53de07a465de8cca0)] Add `getCellsBySheetId` function
    * [[#d23ac45](https://github.com/CSBiology/FsSpreadsheet/commit/d23ac45791de604d83d24b15219852e08e29e0c3)] Replace way to get cells from bySheetIndex to bySheetID
    * [[#3fa594c](https://github.com/CSBiology/FsSpreadsheet/commit/3fa594c84d4fcaaa12f298a13ad962400a97cf89)] Add unit tests for edgecase XlsxFiles
* Bugfixes:
    * [[#99dfa53](https://github.com/CSBiology/FsSpreadsheet/commit/99dfa5329fa02d5e7f358a2eb914f6ac9155d63c)] Fix incorrect casename

### 1.0.0+6f3ddf9 (Released 2023-3-23)
* Additions:
    * [[#2c052d9](https://github.com/CSBiology/FsSpreadsheet/commit/2c052d900bfceafaeb2c76013556e57124505b78)] rename Missing to SheetEntity
    * [[#b3225ef](https://github.com/CSBiology/FsSpreadsheet/commit/b3225efcda3d29f2524ea9306191adeeac66c453)] allow cell, row and column builders yield optional or required operator
    * [[#4a8649a](https://github.com/CSBiology/FsSpreadsheet/commit/4a8649a962cdf6ac02ba83854b893c83d2b000ad)] make SheetEntity type more robust
    * [[#ec31b63](https://github.com/CSBiology/FsSpreadsheet/commit/ec31b63add8228ecd54aacdfe88aa94e5d832595)] change target framework
    * [[#fa28800](https://github.com/CSBiology/FsSpreadsheet/commit/fa2880025dd54aeda84c85e0abe2a2ea17251031)] allow message type to be exception
    * [[#bfd58ae](https://github.com/CSBiology/FsSpreadsheet/commit/bfd58ae0fad619ce68ecf3ab3ce91b51283bc302)] add csv writing capabilities
    * [[#41ea33a](https://github.com/CSBiology/FsSpreadsheet/commit/41ea33af36ac6667a414db66d3bacdc08c6b862d)] replace build script by build project
    * [[#bf9b8ea](https://github.com/CSBiology/FsSpreadsheet/commit/bf9b8ea87760d84d4762fcf7ef069bae1c6c0572)] rename ok operator to some to match SheetEntity case name
    * [[#fd7e783](https://github.com/CSBiology/FsSpreadsheet/commit/fd7e783484f9951a353911d379c14db2c8ce1fd2)] change dsl operator names
    * [[#8e4e4f6](https://github.com/CSBiology/FsSpreadsheet/commit/8e4e4f60f8e265bbe21aff8f67a33f6229bf14ad)] update test project target framework to net7.
    * [[#9ef1e5f](https://github.com/CSBiology/FsSpreadsheet/commit/9ef1e5f813428b095b4c1aea78e45506a19bbc9f)] init restructuring in functional style
    * [[#575f50a](https://github.com/CSBiology/FsSpreadsheet/commit/575f50a84f0bab696e55309619548713bde57b80)] rows now carry cell collection
    * [[#3d7d20f](https://github.com/CSBiology/FsSpreadsheet/commit/3d7d20f97201aebfb6051260aa9e1b1ff753e588)] Add static methods for FsWorksheet & FsRow
    * [[#4fd2d54](https://github.com/CSBiology/FsSpreadsheet/commit/4fd2d5438ab6e7b8270e81fd600aa1e19402229c)] Add functionality & Triple Slash docu
    * [[#7017802](https://github.com/CSBiology/FsSpreadsheet/commit/70178027cf025a15833e85bfe6a058e7d4e4ff33)] Add static methods for FsCell
    * [[#dc4fe4b](https://github.com/CSBiology/FsSpreadsheet/commit/dc4fe4b708c343d0dd4e0a5569d2f0f866d1a203)] Add (static & non-static) methods for FsCellsCollection
    * [[#fd6fe47](https://github.com/CSBiology/FsSpreadsheet/commit/fd6fe477d07b7f15a650ece72d57ede7c7e2cee0)] Format FsRow.fs
    * [[#75d7ca4](https://github.com/CSBiology/FsSpreadsheet/commit/75d7ca4e2a540ed771c5749fd68f1b53c2e80b45)] Add ((non-)static) methods to FsWorkbook.fs
    * [[#eb90e97](https://github.com/CSBiology/FsSpreadsheet/commit/eb90e97fd0a061fa474c2830fe083ec58e78f436)] Add static methods and ///-docu for XLSX writer extensions
    * [[#5666834](https://github.com/CSBiology/FsSpreadsheet/commit/5666834d09f6f0569f483d90a44a7ff43dd01a20)] Add FsTable import functionality
    * [[#0793f35](https://github.com/CSBiology/FsSpreadsheet/commit/0793f354ea98d1232863efdfdedb7b7141447d50)] Move Table.toFsTable to FsExtensions
    * [[#fe41eb2](https://github.com/CSBiology/FsSpreadsheet/commit/fe41eb2c97ae7cb3c7e44956bd7a6fe74f5141e2)] Add more alternative constructor for FsWorksheet
    * [[#e47d58c](https://github.com/CSBiology/FsSpreadsheet/commit/e47d58cedf6d317821e951202c5af474c1fd1812)] Rename OpenXML-based methods
    * [[#352833c](https://github.com/CSBiology/FsSpreadsheet/commit/352833cce4844a48d4373470af068d82dd68cb92)] Comment out unfinished code to prevent building errors
    * [[#0a86811](https://github.com/CSBiology/FsSpreadsheet/commit/0a86811a4a22890af68c76d91699222654d35db6)] Comment out warn-inducing elements
    * [[#3fdee6b](https://github.com/CSBiology/FsSpreadsheet/commit/3fdee6b21f70d652223973af5dc42e62df2d54ba)] Rename FsRangeAdress.fs into correct english
    * [[#a2fc967](https://github.com/CSBiology/FsSpreadsheet/commit/a2fc9677d98fdadb279cd5dd0404c91b9751b0eb)] Add playground script file for testing/developing purposes
    * [[#fe288fd](https://github.com/CSBiology/FsSpreadsheet/commit/fe288fd9518e5d2cfd8eddf253f22dafaba41312)] Rename WorkSheet.fs to match module name
    * [[#b3774e3](https://github.com/CSBiology/FsSpreadsheet/commit/b3774e34cb3195c43c50b161628ddaab458a94cd)] Renamed file again because Git does not understand
    * [[#be9722d](https://github.com/CSBiology/FsSpreadsheet/commit/be9722d8f301ac09d7a6ac232915ae97abbd432e)] Re-renamed file to correct name
    * [[#f766dea](https://github.com/CSBiology/FsSpreadsheet/commit/f766dea964a5b7b0add9cbb36ed63c1811156c83)] Reuse method instead of copypasting 95 % of the code
    * [[#3485e4b](https://github.com/CSBiology/FsSpreadsheet/commit/3485e4b0c717be43a33a89c9e4f1027df8de74a9)] Add annotation
    * [[#1189756](https://github.com/CSBiology/FsSpreadsheet/commit/118975658718869491faa28aaa20d78150784a0e)] Add function to tranlate CellValue to DataType
    * [[#76ae1cf](https://github.com/CSBiology/FsSpreadsheet/commit/76ae1cf818f01ae9bc16f34bad1ea95776e02781)] Format format
    * [[#8710d88](https://github.com/CSBiology/FsSpreadsheet/commit/8710d88eec8994e372d87ba3c9ab601af57a427d)] Add function to get Tables in WorksheetPart
    * [[#ca54bf7](https://github.com/CSBiology/FsSpreadsheet/commit/ca54bf72c2c55f2ac7740978836a3a97d41ef17b)] Play around in playground
    * [[#55b5b61](https://github.com/CSBiology/FsSpreadsheet/commit/55b5b61e27792e955ca04e5345a429f34aa80e30)] Restructure format of FsCellsCollection
    * [[#e83138b](https://github.com/CSBiology/FsSpreadsheet/commit/e83138b03e28a023b89c3cecab9c325ca64e59cd)] Restructure format of FsWorksheet
    * [[#45c5284](https://github.com/CSBiology/FsSpreadsheet/commit/45c528440c4ce122eba922e14c228372d37a160f)] Correct naming mistake
    * [[#290f281](https://github.com/CSBiology/FsSpreadsheet/commit/290f281e13f451a718d94c9659dc9577815235fd)] Add methods to check for FsCell presence
    * [[#80ae5de](https://github.com/CSBiology/FsSpreadsheet/commit/80ae5de4638067090afe9fcf7c1c0ea1f8b0c201)] Change method name for consistency reasons
    * [[#d4e2fd4](https://github.com/CSBiology/FsSpreadsheet/commit/d4e2fd4a4ba8805584fffa5eb463f76b0522e507)] Restructure format of FsRow
    * [[#2d3b618](https://github.com/CSBiology/FsSpreadsheet/commit/2d3b6184978e61ea13747a159f32224f140b2c7d)] Add static methods for retrieving FsCells
    * [[#555cee5](https://github.com/CSBiology/FsSpreadsheet/commit/555cee5e35f5ffcc79243cd6ecb3d298b02587cc)] Add exception for creation with an inappropriate FsCellsCollection
    * [[#3afa531](https://github.com/CSBiology/FsSpreadsheet/commit/3afa5318f30e804c19f61990578e04c45e48e8fa)] Play around in playground
    * [[#9201282](https://github.com/CSBiology/FsSpreadsheet/commit/9201282aaec38beb7d9d2abfb7b679250dcb04d7)] Add critical annotation to properties
    * [[#0abafea](https://github.com/CSBiology/FsSpreadsheet/commit/0abafeab54322ef74808cc8b23f42e762768da9a)] Add method to more easily add FsCells
    * [[#d1b31ad](https://github.com/CSBiology/FsSpreadsheet/commit/d1b31ad868357375ee3c27e09c0b37eb2e00760f)] Play around in playground
    * [[#ea90984](https://github.com/CSBiology/FsSpreadsheet/commit/ea90984b1986408e4c1868ad7c41f13161b3e600)] Add blueprint of method to implement later
    * [[#a004ddc](https://github.com/CSBiology/FsSpreadsheet/commit/a004ddc7b4c7decded6377cf42449900db60e56d)] Add several add methods
    * [[#95a0866](https://github.com/CSBiology/FsSpreadsheet/commit/95a08668a5d678558f26be85974e4dbdbd009dd5)] Add alternate constructor (WIP)
    * [[#efca757](https://github.com/CSBiology/FsSpreadsheet/commit/efca75770f90d6b1482e060aea5502b518b44c4b)] Play around in playground
    * [[#c66189a](https://github.com/CSBiology/FsSpreadsheet/commit/c66189ae8ac013eea370954029a08b4724e6fbb1)] Add Unit Test for DataType
    * [[#910117a](https://github.com/CSBiology/FsSpreadsheet/commit/910117a0808eb012b7b81ccd86dc21d4b18b797f)] Reformat format of FsAddress for readability/findability
    * [[#28b7f54](https://github.com/CSBiology/FsSpreadsheet/commit/28b7f54281de318f57c7c74d3d733e4cdb8d2ed3)] Rename ambiguous method
    * [[#65dbd74](https://github.com/CSBiology/FsSpreadsheet/commit/65dbd746da6106e41ba6143648112947c3bfae8e)] Add methods for FsAddress
    * [[#3337d57](https://github.com/CSBiology/FsSpreadsheet/commit/3337d57f3709bc0a398cee57e1dee24904b10ea3)] Add unit tests for FsAddress
    * [[#ea32c22](https://github.com/CSBiology/FsSpreadsheet/commit/ea32c22bab0b55e0cf06abc843eb033ba867f2bc)] Add more unit tests for FsCell
    * [[#c091bad](https://github.com/CSBiology/FsSpreadsheet/commit/c091bad2f2b9921f7d16a4bf4b0e7a395c3df455)] Let `UpdateIndices` method return the object
    * [[#8526c91](https://github.com/CSBiology/FsSpreadsheet/commit/8526c9100f7d15e1b3855f78991619cc49eb2c38)] Add ignore to not-used value
    * [[#c430d2d](https://github.com/CSBiology/FsSpreadsheet/commit/c430d2d864108c9583b64a52c2a9993249a2c14e)] Play around in playground
    * [[#b4ef99c](https://github.com/CSBiology/FsSpreadsheet/commit/b4ef99cf35db03ae8b14d3c6928a5909f6128592)] Add XML tags to member description
    * [[#5ceff21](https://github.com/CSBiology/FsSpreadsheet/commit/5ceff21dd5451e44011870781eac5b2bb8ffbbeb)] Add XML tags to member description
    * [[#12a5060](https://github.com/CSBiology/FsSpreadsheet/commit/12a50602331a1a5e55661ccc6c00ea55008a2e35)] Sheet Index and SheetId ambiguity
    * [[#dff7243](https://github.com/CSBiology/FsSpreadsheet/commit/dff72435f4d0446cc6daba58aa985434ab6d691c)] Add IO testing data
    * [[#adb5812](https://github.com/CSBiology/FsSpreadsheet/commit/adb58122a564a52088921179a1e882f00d0f1fb7)] Add Test project for FsSpreadsheet.ExcelIO
    * [[#373f774](https://github.com/CSBiology/FsSpreadsheet/commit/373f7748213705fe2fbf6a7d4973d3133688aa87)] Move data folder for Xlsx tests into FsSpreadsheet.ExcelIO folder
    * [[#f354f13](https://github.com/CSBiology/FsSpreadsheet/commit/f354f13f3df8c77330bc4a276422b5f8d0e8b5bc)] Add Expecto test call
    * [[#3f6e431](https://github.com/CSBiology/FsSpreadsheet/commit/3f6e431b511d7284ad06cb8e4472eb47c471125a)] Format format
    * [[#866dc33](https://github.com/CSBiology/FsSpreadsheet/commit/866dc33f339b96079f3017839ad6d5ce3fadaa04)] Add Spreadsheet tests
    * [[#24bdd61](https://github.com/CSBiology/FsSpreadsheet/commit/24bdd6196e1a83ab593fd0e259906caf268c88f9)]  Add Workbook tests
    * [[#0d60c04](https://github.com/CSBiology/FsSpreadsheet/commit/0d60c04fadcefd06d842d2bb93935c2f6792021f)] Add Sheet tests
    * [[#2858ed8](https://github.com/CSBiology/FsSpreadsheet/commit/2858ed83c7f6fa249c26b4349bd5b87fdbeeb1f2)] Add more Spreadsheet tests
    * [[#4c07ac1](https://github.com/CSBiology/FsSpreadsheet/commit/4c07ac1462d079acc1be46ca214af48dfe73957f)] Change tested property
    * [[#77e15a7](https://github.com/CSBiology/FsSpreadsheet/commit/77e15a72f8391fa91ab7aef154d26f27b36867dd)] Add Cell tests
    * [[#f08ba18](https://github.com/CSBiology/FsSpreadsheet/commit/f08ba181c4ba2d0b0173f6ee194145168bc9c52d)] Refactoring FsCell
    * [[#79b4e85](https://github.com/CSBiology/FsSpreadsheet/commit/79b4e85390c24aa3d6ed1aa61d44536432f32e6d)] Merge FsCell Testing
    * [[#653a5af](https://github.com/CSBiology/FsSpreadsheet/commit/653a5af88e12e5d4bb9b841a9e3e228a454b0b36)] Add FsWorksheet add FsCell
    * [[#31b6789](https://github.com/CSBiology/FsSpreadsheet/commit/31b6789055b31e557c0b80382f8d2a2862cf1c69)] Add FsSparseMatrix
    * [[#f518d9f](https://github.com/CSBiology/FsSpreadsheet/commit/f518d9f47d6a8b0a868aa01a6bf3cac383843440)] Add FsCell construtcor IConvertible value
    * [[#31bdfa9](https://github.com/CSBiology/FsSpreadsheet/commit/31bdfa9a8da35a4ffd6ede0a7f6cc618e787c21f)] Update `tryItem` to be 1-based like the OpenXML implementation
    * [[#c1d54b2](https://github.com/CSBiology/FsSpreadsheet/commit/c1d54b2811a220c46cee89142d7afec7e376c750)] Update testCells to have their real values instead of SST index
    * [[#96f5b36](https://github.com/CSBiology/FsSpreadsheet/commit/96f5b36ef94c2ad908e5f9b5ddc0122487cad740)] Add FsWorkbook FromXlsxStream extension
    * [[#939537a](https://github.com/CSBiology/FsSpreadsheet/commit/939537aec618ed3528d24c5448300ee8054f5386)] Refactor FsWorkbook
    * [[#bce8df8](https://github.com/CSBiology/FsSpreadsheet/commit/bce8df8e5c53881ceaa29214232b15eb09832e4d)] Play around in playground
    * [[#38cec72](https://github.com/CSBiology/FsSpreadsheet/commit/38cec72b6cd0321cfa9c1572564e341956b92ed2)] Add ///-XML tags
    * [[#6480c95](https://github.com/CSBiology/FsSpreadsheet/commit/6480c959287eb47c598b6fa1771ad0547dc8d2de)] Add static method for reader method
    * [[#1a066bc](https://github.com/CSBiology/FsSpreadsheet/commit/1a066bc25f60609003d8d5caad705f4b1ce09858)] Add FsWorksheet table methods alongside respective tests
    * [[#80f8aa5](https://github.com/CSBiology/FsSpreadsheet/commit/80f8aa52676575f3ce32541480296030c0bcb4a5)] Git still does not understand changes in letter case, part I
    * [[#ccf81f0](https://github.com/CSBiology/FsSpreadsheet/commit/ccf81f009de1d34e704739a56ebcb59b353d0ad2)] Git still does not understand changes in letter case, part II
    * [[#2368b70](https://github.com/CSBiology/FsSpreadsheet/commit/2368b70fe73937b521d8dc535312e5974f559a90)] Add `.AddTables` functionality (+ unit tests)
    * [[#c08383b](https://github.com/CSBiology/FsSpreadsheet/commit/c08383bf333b856a2a1defae97c524d687afa393)] Reformat code structure
    * [[#c6c064f](https://github.com/CSBiology/FsSpreadsheet/commit/c6c064f470803d40f53f18497b47127fece438d1)] Add methods to get FsWorksheet by name + unit tests
    * [[#b458922](https://github.com/CSBiology/FsSpreadsheet/commit/b458922cf0a2492ed1316b60177aa6b1b2ac361a)] Finalize `FromXlsxStream` method
    * [[#d641f1e](https://github.com/CSBiology/FsSpreadsheet/commit/d641f1edf7c86df54ca70e1c57cc1cec45902c0d)] Add unit tests for `FromXlsxStream` method (WIP)
    * [[#536a1e3](https://github.com/CSBiology/FsSpreadsheet/commit/536a1e366fb08a89c8a0a685baafede5d59fea17)] Add CellValues to DataType conversion (+ unit tests)
    * [[#488b249](https://github.com/CSBiology/FsSpreadsheet/commit/488b249873bbe62470d31ac51b352e36e8606a6b)] Specify member description
    * [[#380a08d](https://github.com/CSBiology/FsSpreadsheet/commit/380a08d92318d4a5425446b2a4046df05bcc1514)] Precise `ofXlsxCell` method alongside with unit tests
    * [[#8e9c768](https://github.com/CSBiology/FsSpreadsheet/commit/8e9c768c255498d8ecef7f412d3f8c204e3ff421)] Update test Xlsx file
    * [[#68f79c5](https://github.com/CSBiology/FsSpreadsheet/commit/68f79c5d6615de403adc75398f468e9895c9795b)] Play around in playground
    * [[#3dc6434](https://github.com/CSBiology/FsSpreadsheet/commit/3dc6434e90b9ed39e11404f2a4ccf303fb7e2929)] Rename incorrectly named source file
    * [[#ba5aa72](https://github.com/CSBiology/FsSpreadsheet/commit/ba5aa7256cb324ea268b0ef891292ca6f9e0e19d)] Finalize unit tests for `fsWorkbookFromStream` method
    * [[#c531ab7](https://github.com/CSBiology/FsSpreadsheet/commit/c531ab76cf3e2b743536635b74be7c1a639e03cb)] Play around in playground
    * [[#393b02d](https://github.com/CSBiology/FsSpreadsheet/commit/393b02d99d90066bdca7ffcd1d5394d1057339c7)] Add `fromXlsxFile` method
    * [[#a76e289](https://github.com/CSBiology/FsSpreadsheet/commit/a76e289199b365e5ab270ac3067cce40246f1369)] Add method for getting FsTables from FsWorkbooks
    * [[#0857107](https://github.com/CSBiology/FsSpreadsheet/commit/085710739aac46b132bffbcb5ff4fab0eb309699)] Add unit tests for `GetTables`
    * [[#0062c62](https://github.com/CSBiology/FsSpreadsheet/commit/0062c629398e24aaa994fd5acd77ea0f714b1a05)] Add method to get the FsWorksheet of an FsTable
    * [[#4999caf](https://github.com/CSBiology/FsSpreadsheet/commit/4999caf8949ea8fcb1471fe28bcabe53b38a0b93)] Add method to add several FsWorksheets + unit tests
    * [[#8ec63bc](https://github.com/CSBiology/FsSpreadsheet/commit/8ec63bc6501ac28cb93256e027a7839388a653bb)] Add some ///-description
    * [[#6cf942d](https://github.com/CSBiology/FsSpreadsheet/commit/6cf942dd17829bdd6d2f56e53b22f40c70d70d83)] Add more FsTable methods
    * [[#a97d858](https://github.com/CSBiology/FsSpreadsheet/commit/a97d858b0c2528a16aede863696b7aec9441c046)] Add address returning methods to FsCellsCollection
    * [[#212e97d](https://github.com/CSBiology/FsSpreadsheet/commit/212e97dc74779f412ad0d64f6ac81a6780dcd000)] Change exception for 0,0 indices in `GetFirstAddress()`
    * [[#e5b112c](https://github.com/CSBiology/FsSpreadsheet/commit/e5b112c876660c525a9e7b22eef2667272763b7b)] Add FsCellsCollection tests (WIP)
    * [[#3147906](https://github.com/CSBiology/FsSpreadsheet/commit/314790689492f369eafa6d4fd921b8b60613c83f)] Add FsTable tests (WIP)
    * [[#f90b0dc](https://github.com/CSBiology/FsSpreadsheet/commit/f90b0dc7fe0ce0731dace8637fd4aab68b37f84a)] Finalize current FsCellsCollection tests
    * [[#8bd24c5](https://github.com/CSBiology/FsSpreadsheet/commit/8bd24c5d9eca64e530492ca68fd1c61f43756681)] Add ///-description
    * [[#2a18dd3](https://github.com/CSBiology/FsSpreadsheet/commit/2a18dd3e203c1b2fcfb4d8bf7de2fb4ff1cad2d2)] Update FsTable tests (WIP)
    * [[#c4291fb](https://github.com/CSBiology/FsSpreadsheet/commit/c4291fba92d7e01c0feb3e2f2a1b53d266f0c7a3)] Add table field unit tests (WIP)
    * [[#ec3aca9](https://github.com/CSBiology/FsSpreadsheet/commit/ec3aca90abb7806388bf12da049a6c6258ecfec8)] Add cellsColl deep copy (WIP)
    * [[#440b213](https://github.com/CSBiology/FsSpreadsheet/commit/440b213e6c93bebbd130182a9fa2c870f17cb28e)] Add rangeBase annotations
    * [[#9b8de2e](https://github.com/CSBiology/FsSpreadsheet/commit/9b8de2ebff1e9ce63d433b73ca3f125bcdfc5739)] Add commented alternative constructor
    * [[#31a9cbf](https://github.com/CSBiology/FsSpreadsheet/commit/31a9cbfc2b9438af05b86e3469c9135332333cff)] Add static methods, update existing methods
    * [[#0e4c2e4](https://github.com/CSBiology/FsSpreadsheet/commit/0e4c2e48eeec3a6e836ce84af9e87afd872554cf)] Update index setter & `DataCells`
    * [[#0764502](https://github.com/CSBiology/FsSpreadsheet/commit/07645029927c7e8d0ebec5a86098ccbd67bad13c)] Finalize FsTableField unit tests
    * [[#f443bfb](https://github.com/CSBiology/FsSpreadsheet/commit/f443bfb6fd37728a563d50cbdc65101341c46065)] Update FsTable functionality (WIP)
    * [[#ee6ef1b](https://github.com/CSBiology/FsSpreadsheet/commit/ee6ef1b9f5d97f71bab26211759c39b081803904)] Play around in playground
    * [[#ead7004](https://github.com/CSBiology/FsSpreadsheet/commit/ead700494f39d4dd8d58671f26120b0e70d740a5)] Add method to add several fields at once to table
    * [[#cfb5e4e](https://github.com/CSBiology/FsSpreadsheet/commit/cfb5e4e2c8f5cef44115fc6b3d537e12f79001f2)] Add unit tests to FsTable
    * [[#df5125c](https://github.com/CSBiology/FsSpreadsheet/commit/df5125ca14a5157c725a29841f1dbff069a000fe)] Add method for creating range cols out of addresses
    * [[#43aad4a](https://github.com/CSBiology/FsSpreadsheet/commit/43aad4a4b5f4ffd76cf1d7b64ae669417240637b)] Adjust method return for FsAddress
    * [[#6966580](https://github.com/CSBiology/FsSpreadsheet/commit/6966580264237f37712cf5a5c7fb8ffed1ca2b21)] Reformat FsCell structure
    * [[#291d6c8](https://github.com/CSBiology/FsSpreadsheet/commit/291d6c816cbb9f8c2eb984402ce20c91654c0293)] Setup docs, add some methods, downgrade to .NET 6, ...
    * [[#772a315](https://github.com/CSBiology/FsSpreadsheet/commit/772a3152c8c615b6627d21dad7fa3443fbeee790)] Add build.sh to sln
    * [[#811f491](https://github.com/CSBiology/FsSpreadsheet/commit/811f491c52b72ae5d54370ef0babf1e781b87b5c)] Update build and deploy actions
    * [[#cbc91db](https://github.com/CSBiology/FsSpreadsheet/commit/cbc91db43c2a09bb5b16e0bd7081af4ca4c914dc)] Update Logo
    * [[#af5a98b](https://github.com/CSBiology/FsSpreadsheet/commit/af5a98baedff5ff2c385e373b4bdb5ece555e18f)] Add FsTable unit tests (WIP)
    * [[#2069ed9](https://github.com/CSBiology/FsSpreadsheet/commit/2069ed947aa1dd5c7c7a76a8c7c053b7c408996a)] Add deep copy methods to all relevant classes
* Deletions:
    * [[#89c3945](https://github.com/CSBiology/FsSpreadsheet/commit/89c39457eac24def323c36a50d68b7fc04295670)] Add remove methods to FsCellsCollection
    * [[#8647c3e](https://github.com/CSBiology/FsSpreadsheet/commit/8647c3e7a06edeaf61699fffa0c8c956a915d049)] Add methods to remove FsCell values
    * [[#153c485](https://github.com/CSBiology/FsSpreadsheet/commit/153c485c9305b8e21e8b545e9299fd3e3f3608fc)] Delete sample Unit Test
    * [[#7269a09](https://github.com/CSBiology/FsSpreadsheet/commit/7269a09dcfd91734d27592016613bb54061409e3)] Delete wrong exception
    * [[#55a766a](https://github.com/CSBiology/FsSpreadsheet/commit/55a766abe4d163a7eccec61d9baa28bc4362e7fe)] Delete '.' from message due to it confusing YoloDev adapter
    * [[#cd32696](https://github.com/CSBiology/FsSpreadsheet/commit/cd326967f2ab76bdd065cd3293a6ab5c7f646bf1)] FsCell Remove alternatve constructors
* Bugfixes:
    * [[#94211fc](https://github.com/CSBiology/FsSpreadsheet/commit/94211fca8cbf3fb90883c8fdfea1086517f4d1fb)] Fix error (wrong member call)
    * [[#810424f](https://github.com/CSBiology/FsSpreadsheet/commit/810424fadaf432587c12e24d550a30578ed7598c)] Fix critical FsWorksheet error and add non-static methods for functionalities
    * [[#5a8fff4](https://github.com/CSBiology/FsSpreadsheet/commit/5a8fff4a078d732cfb99bcf4caaf5b4c6e8970fb)] Fix naming-related error
    * [[#5b857a6](https://github.com/CSBiology/FsSpreadsheet/commit/5b857a67d69181be26602bcc9064f5782fa0553d)] Fix critical error in getter method
    * [[#ef97805](https://github.com/CSBiology/FsSpreadsheet/commit/ef97805b39713271804a6c168bfa813738e1cce7)] Fix critical error according to the correct meaning of properties
    * [[#122964b](https://github.com/CSBiology/FsSpreadsheet/commit/122964b6f95b441ef23470e30aa8eaf47615a5db)] Fix merge
    * [[#5fd9bc2](https://github.com/CSBiology/FsSpreadsheet/commit/5fd9bc24a29c6be195cdea24044547aa5a24c5d8)] Fix error when trying to get colIndex
    * [[#7bf051d](https://github.com/CSBiology/FsSpreadsheet/commit/7bf051d1316e972220f5bd37cb2ae9f4cae1ebfb)] Fix typo in test file
    * [[#afdab8e](https://github.com/CSBiology/FsSpreadsheet/commit/afdab8ec7eddf13e6dc4c1346251959db8b329d2)] Fix critical XlsxCell error (+ respective unit test)
    * [[#05437d7](https://github.com/CSBiology/FsSpreadsheet/commit/05437d7b66b72d86e51a0cf58b3605f7537e1e0c)] Fix null reference error (+ unit test)
    * [[#d6abf25](https://github.com/CSBiology/FsSpreadsheet/commit/d6abf25c2a1fd8b4711d01c12e853a798d57e4ff)] Fix possible error due to hard coding
    * [[#1bf0ffd](https://github.com/CSBiology/FsSpreadsheet/commit/1bf0ffd2900e94a7d7fbb761304b05fc9146de53)] Fix typo
    * [[#448ccda](https://github.com/CSBiology/FsSpreadsheet/commit/448ccda2ae0ff384b218f1a7d558d1d0545ff449)] Add and fix table field functionality
    * [[#5437128](https://github.com/CSBiology/FsSpreadsheet/commit/5437128d01655a8f26b8d17a16a07e5aeaf26d02)] Finalize FsTable, fix indexing error
    * [[#dceac83](https://github.com/CSBiology/FsSpreadsheet/commit/dceac832d0ef8c5656265744fe30c56c07dff8b9)] Update table unit tests, fix XML comment typos
    * [[#6f3ddf9](https://github.com/CSBiology/FsSpreadsheet/commit/6f3ddf9120e9b5fc001ba43bb4b0bbebeab7d4ed)] Fix critical return type error

### 0.2.2+b86ecfd (Released 2023-2-3)
* Additions:
    * [[#b86ecfd](https://github.com/CSBiology/FsSpreadsheet/commit/b86ecfda1831d8793df9d49060787901916fcca6)] add additional yield method to DSL

### 0.2.1+927f365 (Released 2022-12-12)
* Additions:
    * latest commit #927f365
* Bugfixes:
    * [[#927f365](https://github.com/CSBiology/FsSpreadsheet/commit/927f3652a8d2b71ad7930f87c486ce606cb2d80e)] fix spacepreservation

### 0.2.0+8084a32 (Released 2022-12-12)
* Additions:
    * [[#8084a32](https://github.com/CSBiology/FsSpreadsheet/commit/8084a32ef40e54a0f38c6bac67246d5f8deb9458)] allow for trailing spaces to be preserved when writing strings

### 0.1.7+a227766 (Released 2022-6-24)
* Additions:
    * latest commit #a227766
* Bugfixes:
    * [[#a227766](https://github.com/CSBiology/FsSpreadsheet/commit/a2277668b273ebb3f6d556de148cd7b176d61dbd)] fix table id determination

### 0.1.6+807ec9b (Released 2022-6-21)
* Additions:
    * [[#d097bd3](https://github.com/CSBiology/FsSpreadsheet/commit/d097bd37bf2b742af0465a428abeea2b4000a49a)] add table to dsl
    * [[#cfc60cc](https://github.com/CSBiology/FsSpreadsheet/commit/cfc60cca3eaf9235b34202f5deb2230e2edc8553)] start working on TableBuilder for DSL
    * [[#a85093f](https://github.com/CSBiology/FsSpreadsheet/commit/a85093f23613c11fb8131e0adbee6124ff744b9a)] Add basic coding examples to ReadMe.md

### 0.1.5+7250efd (Released 2022-5-10)
* Additions:
    * latest commit #7250efd
* Bugfixes:
    * [[#7250efd](https://github.com/CSBiology/FsSpreadsheet/commit/7250efdd8dbc8fc95fb43050ea3aeb3070dabae2)] fix dataType of cells always being text
    * [[#6d12edf](https://github.com/CSBiology/FsSpreadsheet/commit/6d12edfa6520349f29e74f4be2d1d043049cc93d)] fix bug where empty rows made Excel writer crash
    * [[#b70044a](https://github.com/CSBiology/FsSpreadsheet/commit/b70044a45e5296597c032f09bf106b91d1139c16)] fix dsl section order

### 0.1.4+4f2bb3b (Released 2022-5-9)
* Additions:
    * latest commit #4f2bb3b
    * [[#4f2bb3b](https://github.com/CSBiology/FsSpreadsheet/commit/4f2bb3bac5b84bc04f8ddf99accc34eccb877870)] allow ColumnBuilder to yield Missing values

### 0.1.3+a18c48b (Released 2022-4-12)
* Additions:
    * latest commit #a18c48b
    * [[#5bf131f](https://github.com/CSBiology/FsSpreadsheet/commit/5bf131f2812d57f4bc4e0a91e030b8d6c573e555)] add convenience functions for dropping dsl items
    * [[#52b95e0](https://github.com/CSBiology/FsSpreadsheet/commit/52b95e038bac177aa31d7e86e7dd7e4f5e1aa8ee)] add cellbuilder
    * [[#705edd9](https://github.com/CSBiology/FsSpreadsheet/commit/705edd9e773fb6e7684105075b30f8930ba732d5)] rework dsl operators

### 0.1.2+c854ed6 (Released 2022-4-11)
* Additions:
    * latest commit #c854ed6
* Bugfixes:
    * [[#9f2903d](https://github.com/CSBiology/FsSpreadsheet/commit/9f2903dde2675de6b7755713248e5106f9651048)] fix for loops in dsl

### 0.1.1+aae6edd (Released 2022-3-31)
* Additions:
    * latest commit #aae6edd
* Bugfixes:
    * [[#aae6edd](https://github.com/CSBiology/FsSpreadsheet/commit/aae6edd8981cad33fc5ad374c824f065c8adb288)] fix row functions failing when reading empty cells

### 0.1.0+268b900 (Released 2022-3-31)
* Additions:
    * latest commit #268b900
    * [[#268b900](https://github.com/CSBiology/FsSpreadsheet/commit/268b90012e90a038ba9bfbad6ef4ff6d8f3ebda4)] add columnbuilder tp dsl
    * [[#eed6806](https://github.com/CSBiology/FsSpreadsheet/commit/eed680645d473e51a3bb47b9d9db20d0e8eb5f6b)] add sheet and workbook builder to DSL
    * [[#ed45200](https://github.com/CSBiology/FsSpreadsheet/commit/ed452001d696457e881d9a5bd390fb58455fa3de)] add optional and required cells to DSL

### 0.0.2+2f09131 (Released 2022-1-21)
* Additions:
    * latest commit #2f09131
    * [[#2f09131](https://github.com/CSBiology/FsSpreadsheet/commit/2f091315c3a587c811d44784ed92be9a6b8fad74)] add basic dsl

### 0.0.1+7be5a42 (Released 2022-1-20)
* Additions:
    * Setup basic types based on ClosedXML
    * Add Sheetbuilder based on ClosedXML.SimpleSheets
    * Copy DocumentFormat.OpenXML wrapper functions from FSharpSpreadsheetML
    * Setup excel write functionality

