
open System

let rec getLines lines :String List  =
    let a = Console.ReadLine()
    sprintf "%s" a
    getLines (List.append lines [a])

let a: String list = getLines []

a |> List.iter (fun x -> printf "%s " x)