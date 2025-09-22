module FsSpreadsheet.HashCodes

let mergeHashes (hash1 : int) (hash2 : int) : int =
    0x9e3779b9 + hash2 + (hash1 <<< 6) + (hash1 >>> 2)

let hashDateTime (dt : System.DateTime) : int =
    let mutable acc = 0
    acc <- mergeHashes acc dt.Year
    acc <- mergeHashes acc dt.Month
    acc <- mergeHashes acc dt.Day
    acc <- mergeHashes acc dt.Hour
    acc <- mergeHashes acc dt.Minute
    acc <- mergeHashes acc dt.Second
    acc
    

let hash obj =
    obj.GetHashCode()

let boxHashOption (a: 'a option) : obj =
    if a.IsSome then a.Value.GetHashCode() else (0).GetHashCode()
    |> box

let boxHashArray (a: 'a []) : obj =
    a 
    // from https://stackoverflow.com/a/53507559
    |> Array.fold (fun acc o -> 
        hash o
        |> mergeHashes acc) 0
    |> box

let boxHashSeq (a: seq<'a>) : obj =
    a 
    // from https://stackoverflow.com/a/53507559
    |> Seq.fold (fun acc o -> 
        hash o
        |> mergeHashes acc) 0
    |> box