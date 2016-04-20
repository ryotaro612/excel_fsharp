open System
open System.IO
let readLines filePath = File.ReadLines(filePath)

let lines = readLines "ramdom.csv"

for word in lines do 
  printfn "word: %s" word


open System.Text.RegularExpressions

// 
let (|Regex|_|) pattern str =
    let r = Regex.Match(str, pattern)
    if r.Success then Some r.Value
    else None

let printMatch input =
    match input with
    | Regex "abc" s -> printfn "match: %s" s
    | Regex "def" s -> printfn "match: %s" s
    | _ -> printfn "none"

printMatch "abcdef"
printMatch "alter"
printMatch "fooba"