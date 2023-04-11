module FsRangeAddressTests

open Expecto
open FsSpreadsheet


let dummyRa1 = FsRangeAddress("A1:B1")
let dummyRa2 = FsRangeAddress("D1:E1")
let dummyRa3 = FsRangeAddress("B1:C1")
let dummyRa4 = FsRangeAddress("A1:A5")
let dummyRa5 = FsRangeAddress("A2:A5")
let dummyRa6 = FsRangeAddress("B2:D4")
let dummyRa7 = FsRangeAddress("C3:E17")
let dummyRa8 = FsRangeAddress("D5:E17")


[<Tests>]
let fsRangeAddressTests =
    testList "FsRangeAddress" [
        testList "Overlaps" [
            testCase "Returns false if not overlapping" <| fun _ ->
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa2 |> not) "Returned true though A1:B1 and D1:E1 don't overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa2 dummyRa3 |> not) "Returned true though D1:E1 and B1:C1 don't overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa2 dummyRa4 |> not) "Returned true though D1:E1 and A1:A5 don't overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa5 |> not) "Returned true though A1:B1 and A2:A5 don't overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa6 |> not) "Returned true though A1:B1 and B2:D4 don't overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa8 |> not) "Returned true though A1:B1 and D5:E17 don't overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa6 dummyRa8 |> not) "Returned true though B2:D4 and D5:E17 don't overlap"
            testCase "Returns true if overlapping" <| fun _ ->
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa1) "Returned false though A1:B1 and A1:B1 overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa3) "Returned false though A1:B1 and B1:C1 overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa1 dummyRa4) "Returned false though A1:B1 and A1:A5 overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa6 dummyRa7) "Returned false though B2:D4 and C3:E17 overlap"
                Expect.isTrue (FsRangeAddress.overlaps dummyRa7 dummyRa8) "Returned false though C3:E17 and D5:E17 overlap"
        ]
    ]