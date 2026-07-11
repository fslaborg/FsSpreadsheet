module FsSpreadsheet.CrossAsync

open Fable.Core


// We need the `<'T>` here as it allows to mark the function/variable inline for better output.
let inline crossAsync<'T> =
#if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    promise
#else
    async
#endif

type CrossAsync<'T> =
#if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    JS.Promise<'T>
#else
    Async<'T>
#endif

let startSequential (starterF : 'T -> CrossAsync<'U>) (tasks : 'T seq) : CrossAsync<'U []> =
    let rec loop (en : System.Collections.Generic.IEnumerator<'T>) =
        crossAsync {
            if en.MoveNext() then
                let! r = starterF en.Current
                let! following = loop en
                return Array.append [|r|] following
            else return [||]
        }
    loop (tasks.GetEnumerator())

let all (tasks : CrossAsync<'T> seq) : CrossAsync<'T []> =
    #if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    Promise.all tasks
    #else
    Async.Sequential tasks
    #endif

let map f v =
    crossAsync {
        let! v = v
        return f v
    }

let asAsync (v:CrossAsync<'T>) =
    #if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    Async.AwaitPromise v
    #else
    v
    #endif

let fromPromise (v:JS.Promise<'T>) =
    #if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    v
    #else
    Async.AwaitPromise v
    #endif

// Promise rejections in JS are often plain objects and not proper Exceptions.
// Normalize them early so test failures show actionable messages.
[<Emit("(function (e) { if (e == null) return 'Promise rejected with null/undefined'; if (typeof e === 'string') return e; if (e.stack) return e.stack; if (e.message) return e.message; try { return JSON.stringify(e); } catch (_) { return String(e); } })($0)")>]
let private jsErrorToString (e: obj) : string = jsNative

let errorToString (e: obj) : string =
    #if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    jsErrorToString e
    #else
    match e with
    | :? exn as ex -> ex.ToString()
    | _ -> string e
    #endif

let catchWith (f : exn -> 'T) (p : CrossAsync<'T>) : CrossAsync<'T> =
    #if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    Promise.catch f p
    #else
    async {
        let! r = Async.Catch p
        match r with
        | Choice1Of2 x -> return x
        | Choice2Of2 e -> return f e
    }
    #endif
    
let catchAsResult (p : CrossAsync<'T>) : CrossAsync<Result<'T,string>>=
    #if FABLE_COMPILER_JAVASCRIPT || FABLE_COMPILER_TYPESCRIPT
    p
    |> Promise.map (fun x -> Ok x)
    |> Promise.catch (fun e -> Error (errorToString e))
    #else
    async {
        let! r = Async.Catch p
        match r with
        | Choice1Of2 x -> return Ok x
        | Choice2Of2 e -> return Error (e.ToString())
    }
    #endif