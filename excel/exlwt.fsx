open System

let rec getLines lines :String List  =
    match Console.ReadLine() with
    | null -> lines
    | str -> getLines (List.append lines [str])

let a: String list = getLines []

a |> List.iter (fun l -> printfn "%s" l)