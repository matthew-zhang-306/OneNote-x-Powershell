try {
    cd $psISE.CurrentFile.FullPath.Replace("\Application.ps1","")
    
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

    Add-Type -AssemblyName Microsoft.Office.Interop.OneNote
    Add-Type -Path "WinSCP/WinSCPnet.dll"
}
finally {
    & ".\OneNote.ps1"
}