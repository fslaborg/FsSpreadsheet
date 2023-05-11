module DSL.CellBuilder


#if FABLE_COMPILER
open Fable.Mocha
#else
open Expecto
#endif
open FsSpreadsheet
open FsSpreadsheet.DSL


let main = 
    testList "DSL.CellBuilder" [
        testCase "simple int" <| fun _ ->
            let cell = 
                cell {
                    1
                }
            let expected = Value(DataType.Number,"1"),None
            Expect.equal cell expected "Cell differs"
        testCase "simple string" <| fun _ ->
            let cell = 
                cell {
                    "Hello"
                }
            let expected = Value(DataType.String,"Hello"),None
            Expect.equal cell expected "Cell differs"
        testCase "concated string" <| fun _ ->
            let cell = 
                cell {
                    Concat ';'
                    "Hello" 
                    " " 
                    "World"
                }
            let expected = Value(DataType.String,"Hello; ;World"),None
            Expect.equal cell expected "Cell differs"
    ]

    