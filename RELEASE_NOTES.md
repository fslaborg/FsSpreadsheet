### 1.0.0+6f3ddf9 (Released 2023-3-23)
* Additions:
    * [[#2c052d9](https://github.com/CSBiology/FsSpreadsheet/commit/2c052d900bfceafaeb2c76013556e57124505b78)] rename Missing to SheetEntity
    * [[#b3225ef](https://github.com/CSBiology/FsSpreadsheet/commit/b3225efcda3d29f2524ea9306191adeeac66c453)] allow cell, row and column builders yield optional or required operator
    * [[#4a8649a](https://github.com/CSBiology/FsSpreadsheet/commit/4a8649a962cdf6ac02ba83854b893c83d2b000ad)] make SheetEntity type more robust
    * [[#ec31b63](https://github.com/CSBiology/FsSpreadsheet/commit/ec31b63add8228ecd54aacdfe88aa94e5d832595)] change target framework
    * [[#7b777d6](https://github.com/CSBiology/FsSpreadsheet/commit/7b777d6bd2464da401b20ec9a6fa6df668d1ff2a)] Merge pull request #8 from CSBiology/DSL
    * [[#fa28800](https://github.com/CSBiology/FsSpreadsheet/commit/fa2880025dd54aeda84c85e0abe2a2ea17251031)] allow message type to be exception
    * [[#81891f4](https://github.com/CSBiology/FsSpreadsheet/commit/81891f4456c2aa3b108ee72a04ab8bce2a8ba487)] Merge pull request #9 from CSBiology/DSL
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
    * [[#8291561](https://github.com/CSBiology/FsSpreadsheet/commit/8291561f147c46bb87928c6754e22e85c4e02fd4)] Merge pull request #17 from CSBiology/datamodel
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
    * latest commit #b86ecfd
    * [[#b86ecfd](https://github.com/CSBiology/FsSpreadsheet/commit/b86ecfda1831d8793df9d49060787901916fcca6)] add additional yield method to DSL

### 0.2.1+927f365 (Released 2022-12-12)
* Additions:
    * latest commit #927f365
* Bugfixes:
    * [[#927f365](https://github.com/CSBiology/FsSpreadsheet/commit/927f3652a8d2b71ad7930f87c486ce606cb2d80e)] fix spacepreservation

### 0.2.0+8084a32 (Released 2022-12-12)
* Additions:
    * latest commit #8084a32
    * [[#8084a32](https://github.com/CSBiology/FsSpreadsheet/commit/8084a32ef40e54a0f38c6bac67246d5f8deb9458)] allow for trailing spaces to be preserved when writing strings

### 0.1.7+a227766 (Released 2022-6-24)
* Additions:
    * latest commit #a227766
* Bugfixes:
    * [[#a227766](https://github.com/CSBiology/FsSpreadsheet/commit/a2277668b273ebb3f6d556de148cd7b176d61dbd)] fix table id determination

### 0.1.6+807ec9b (Released 2022-6-21)
* Additions:
    * latest commit #807ec9b
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
    * latest commit #7be5a42
    * Setup basic types based on ClosedXML
    * Add Sheetbuilder based on ClosedXML.SimpleSheets
    * Copy DocumentFormat.OpenXML wrapper functions from FSharpSpreadsheetML
    * Setup excel write functionality

