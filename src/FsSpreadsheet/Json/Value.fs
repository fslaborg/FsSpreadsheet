module FsSpreadsheet.Json.Value

open Thoth.Json.Core
open FsSpreadsheet

#if FABLE_COMPILER_PYTHON
module PyTime = 

    open Fable.Core
    open Fable.Core.PyInterop

    // Currently in Fable, a created datetime object will contain a timezone. This is not allowed in python xlsx, so we need to remove it.
    // Unfortunately, the timezone object in python is read-only, so we need to create a new datetime object without timezone.
    // For this, we use the fromtimestamp method of the datetime module and convert the timestamp to a new datetime object without timezone.

    type DateTimeStatic =
        [<Emit("$0.fromtimestamp(timestamp=$1)")>]
        abstract member fromTimeStamp: timestamp:float -> System.DateTime

    [<Import("datetime", "datetime")>]
    let DateTime : DateTimeStatic = nativeOnly

    let toUniversalTimePy (dt:System.DateTime) = 
        
        dt.ToUniversalTime()?timestamp()
        |> DateTime.fromTimeStamp
#endif





module Decode =

    let datetime: Decoder<System.DateTime> =
        #if FABLE_COMPILER_PYTHON
        Decode.datetimeLocal |> Decode.map PyTime.toUniversalTimePy
        #else
        { new Decoder<System.DateTime> with
            member _.Decode(helpers, value) =
                if helpers.isString value then
                    match System.DateTime.TryParse(helpers.asString value) with
                    | true, datetime -> datetime |> Ok
                    | _ -> ("", BadPrimitive("a datetime", value)) |> Error
                else
                    ("", BadPrimitive("a datetime", value)) |> Error
        }
        #endif

let encode (value : obj) =
    match value with
    | :? string as s -> Encode.string s
    | :? float as f -> Encode.float f
    | :? int as i -> Encode.int i
    | :? bool as b -> Encode.bool b
    | :? System.DateTime as d -> 
        d.ToString("O", System.Globalization.CultureInfo.InvariantCulture).Split('+').[0]
        |> Encode.string
    | _ -> Encode.nil



let decode =    
    Decode.oneOf [
        Decode.bool |> Decode.map (fun b -> b :> obj, DataType.Boolean)
        Decode.int |> Decode.map (fun i -> i :> obj, DataType.Number)
        Decode.float |> Decode.map (fun f -> f :> obj, DataType.Number)
        Decode.datetime |> Decode.map (fun d -> d :> obj, DataType.Date)
        Decode.string |> Decode.map (fun s -> s :> obj, DataType.String)
    ]