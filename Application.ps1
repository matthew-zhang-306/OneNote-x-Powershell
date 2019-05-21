Add-Type -AssemblyName Microsoft.Office.Interop.OneNote

cd $psISE.CurrentFile.FullPath.Replace("\Application.ps1","")

& ".\OneNote.ps1"