using namespace Microsoft.Office.Interop
using namespace System.Xml

Clear-Host

try {
    $onenote = New-Object -ComObject OneNote.Application
    [xml]$hierarchy = ""
    $onenote.GetHierarchy("", [OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)


    [XmlElement]$notebook = $null
    foreach ($notebookXml in $hierarchy.Notebooks.Notebook) {
        if ($notebookXml.Name.Contains("Sai")) {
            $notebook = $notebookXml
        }
    }

    $fromgroup = $null
    $togroup = $null
    foreach ($sectiongroup in $notebook.SectionGroup) {
        if ($sectiongroup.Name.Contains("Thursday")) {
            $fromgroup = $sectiongroup
        }
        elseif ($sectiongroup.Name.Contains("Friday")) {
            $togroup = $sectiongroup
        }
    }


    $tosection = $null
    foreach ($section in $togroup.Section) {
        if ($section.Name.Contains("Math")) {
            $tosection = $section
        }
    }


    $targetpage = $null
    $targetcontent = $null
    foreach ($page in $fromgroup.Section.Page) {
        if ($page.Name.Contains("to be moved")) {
            $targetpage = $page
            $onenote.GetPageContent($page.ID, [ref]$targetcontent)
        }
    }


    $newpageid = $null
    $onenote.CreateNewPage($tosection.ID, [ref]$newpageid)
    $targetcontent = $targetcontent.Replace($targetpage.ID, $newpageid)


    Write-Host $targetcontent


    $onenote.UpdatePageContent($targetcontent)
    $onenote.DeleteHierarchy($targetpage.ID)
}
catch {
    $_
}