try {
    if ($PSScriptRoot -ne $null) {
        cd $PSScriptRoot
    } else {
        cd $psISE.CurrentFile.FullPath.Replace("\Application.ps1","")
    }

    # Execution policy might not work because of the "more specific scope" error so this try/catch block works around it
    try {
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
    }
    catch {
        "Could not set execution policy: "
        $_
        Start-Sleep -s 3
    }

    Add-Type -AssemblyName Microsoft.Office.Interop.OneNote
    Add-Type -Path "WinSCP/WinSCPnet.dll"

    & ".\OneNote.ps1"
}
catch {
    "Error running: "
    $_
    Start-Sleep -s 15
}