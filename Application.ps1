$PSScriptRoot = $psISE.CurrentFile.FullPath.Replace("\Application.ps1","")
cd $PSScriptRoot
$env:PSModulePath = $PSScriptRoot

cls
Add-Type -AssemblyName Microsoft.Office.Interop.OneNote

& ".\Rectangle.ps1"
& ".\DateHelper.ps1"
& ".\HtmlCreator.ps1"
& ".\Ink.ps1"
& ".\Image.ps1"
& ".\Page.ps1"
& ".\PageHtml.ps1"
& ".\Section.ps1"
& ".\Notebook.ps1"
& ".\OneNote.ps1"

