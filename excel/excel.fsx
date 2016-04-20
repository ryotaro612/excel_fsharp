#r "Microsoft.Office.Interop.Excel"
(* Reference: https://msdn.microsoft.com/EN-US/library/microsoft.office.interop.excel.aspx *)
open System
open System.IO
open System.Reflection
open Microsoft.Office.Interop.Excel

// If Visible is false, you can't see the following operations.
let app = ApplicationClass(Visible=true)
let wb = app.Workbooks.Add()
let sheet = wb.Worksheets.[1] :?> _Worksheet
sheet.Name <- "add a sheet"
sheet.Copy(sheet, Type.Missing) // copy a sheet.
// wb.SaveAs("abc.xlsx")

sheet.Cells.[1,1] <- "foo" // add value

(* The value of a range can be set using Value2 property.
   https://msdn.microsoft.com/en-us/library/hh297098(v=vs.100).aspx *)
// Store data in arrays of strings or floats
let rnd = new Random()
let titles = [| "No"; "Maybe"; "Yes" |]
let names = Array2D.init 10 1 (fun i _ -> string('A' + char(i)))
let data = Array2D.init 10 3 (fun _ _ -> rnd.NextDouble())

// Populate data into Excel worksheet
sheet.Range("C2", "E2").Value2 <- titles
sheet.Range("B3", "B12").Value2 <- names
sheet.Range("C3", "E12").Value2 <- data

let charts : ChartObjects = sheet.ChartObjects() :?> ChartObjects

