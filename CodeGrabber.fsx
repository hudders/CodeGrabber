#r @"..\..\_ref\FSharp.Data.dll"
#r @"..\..\_ref\WebDriver.dll"
#r @"..\..\_ref\Newtonsoft.Json.dll"
#r @"..\..\_ref\SizSelCsZzz.dll"
#r @"..\..\_ref\canopy.dll"

#r "System.Xml.dll"
#r "System.Xml.Linq.dll"
#r "Microsoft.Office.Interop.Excel"
#r "office"

open System
open System.IO

open canopy
open System.IO
open System.Xml
open System.Xml.Linq
open Microsoft.Office.Interop

let convertPath (path : string, extension : string) =
    let fileList = Seq.toList (System.IO.Directory.EnumerateFiles(path, "*." + extension))
    fileList.[0]

let xmlPath = convertPath(@"C:\x_FSharpStuff\_dat\lookup\", "xml")
let xlsPath = convertPath(@"C:\x_FSharpStuff\_dat\lookup\", "xlsx")

let xlApp = new Excel.ApplicationClass()
xlApp.Visible <- true
let xlsFile = (xlApp.Workbooks.Open(xlsPath))

let xlsTab (xls : Excel.Workbook, tab : string) = xls.Worksheets.[tab] :?> Excel.Worksheet

let clearDown(xlsFile : Excel.Worksheet) =
    let rec clean n =
        let cell1 = "E" + n.ToString()
        let cell2 = "F" + n.ToString()
        if n < 1001 then
            xlsFile.Range(cell1, cell2).Value2 <- null
            clean(n + 1)
    clean 3

let enterValue(xlsFile : Excel.Worksheet, value : string, column : string) =
    let total = column + "1000"
    let value = value.Replace("&amp;","&")
    let rec loop n =
        let cell = column + (n.ToString())
        if xlsFile.Range(cell, cell).Value2 = null then
            xlsFile.Range(cell, cell).Value2 <- value
            xlsFile.Range(total, total).Value2 <- n
            printf "."
        else
            loop (n + 1)
    if xlsFile.Range(total, total).Value2 = null || xlsFile.Range(total, total).Value2.ToString() = "" then
        loop 3
    else
        loop (System.Convert.ToInt32(xlsFile.Range(total, total).Value2.ToString()))

let grabXml(xlsFile : Excel.Worksheet, nodePath : string, column : string) =
    if File.Exists(xmlPath) && File.Exists(xlsPath) then
        printfn " "
        printf "Working"
        let xml = XDocument.Load(xmlPath).ToString()
        let doc = new XmlDocument() in doc.LoadXml xml
        doc.SelectNodes nodePath
            |> Seq.cast<XmlNode>
            |> Seq.iter (fun node -> enterValue (xlsFile, node.InnerXml.ToString(), column))
    else
        printfn "Something doesn't exist - are your xml and excel files in the right places?"

// Obviously we will need some logic to decide which paths to use (i.e. for occupations, modifications, etc)
// For the moment this just works for pet breeds...

clearDown(xlsTab(xlsFile, "Breed Codes"))
grabXml(xlsTab(xlsFile, "Breed Codes"), "//lookup/breeds/breed/ctmvalue", "E")
grabXml(xlsTab(xlsFile, "Breed Codes"), "//lookup/breeds/breed/providervalue", "F")