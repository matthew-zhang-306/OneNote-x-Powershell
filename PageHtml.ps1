using namespace System.Collections.Generic
using namespace System.Xml

class PageHtml {
    [string]$NotebookName
    [string]$SectionName
    [string]$PageName
    [string]$Tag

    PageHtml([Page]$page) {
        if ($page -ne $null) {
            $this.NotebookName = $page.Section.Notebook.Name
            $this.SectionName = $page.Section.Name
            $this.PageName = $page.Name
            $this.Tag = $page.TagName
        }
    }
}