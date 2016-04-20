#r "Microsoft.Office.Interop.Excel"

open System
open System.IO
open System.Reflection
open Microsoft.Office.Interop.Excel

(* let app = Application(Visible=true)*)
let app = ApplicationClass(Visible=true)
let wb = app.Workbooks.Add()
let sheet = wb.Worksheets.[1] :?> _Worksheet
wb.SaveAs("abc.xlsx")
