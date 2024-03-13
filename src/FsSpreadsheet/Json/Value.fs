module FsSpreadsheet.Json.Value

open Thoth.Json.Core

let encode (value : obj) =
    match value with
    | :? string as s -> Encode.string s
    | :? int as i -> Encode.int i
    | :? float as f -> Encode.float f
    | :? bool as b -> Encode.bool b
    | _ -> Encode.nil

let decode : Decoder<obj> =
    Decode.oneOf [
        Decode.string |> Decode.map (fun s -> s :> obj)
        Decode.int |> Decode.map (fun i -> i :> obj)
        Decode.float |> Decode.map (fun f -> f :> obj)
        Decode.bool |> Decode.map (fun b -> b :> obj)
    ]