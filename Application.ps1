[bool]$didWork = $false

try {
    if ($PSScriptRoot -ne $null) {
        Set-Location $PSScriptRoot
    } else {
        Set-Location $psISE.CurrentFile.FullPath.Replace("\Application.ps1","")
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

    $didWork = $true
}
catch {
    "Error running: "
    $_
    Start-Sleep -s 15
}


if ($didWork) {
    & ".\OneNote.ps1"
}