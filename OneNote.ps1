using namespace Microsoft.Office.Interop
using namespace System.Collections.Generic
using namespace System.Xml

# Magic word clears the console
cls

# If Powershell for some reason doesn't recognize OneNote classes, type this into the console to magically fix it
Add-Type -AssemblyName Microsoft.Office.Interop.OneNote


# To be reworked into HtmlCreator in a separate file
class HtmlManager {
    static [string]$Style
    
    static [string]MakeTitle([string]$name) {
        return "<title>" + $name + "</title>"
    }
    static [string]GetFullHead([string]$title) {
        return [HtmlManager]::MakeTitle($title) + [HtmlManager]::Style
    }

    static [string]MakePre([string]$name) {
        return "<h1>" + $name + "</h1>"
    }
}

# Set css html (done outside of the classes because the syntax breaks when inserting leading tabs)
[HtmlManager]::Style = @"
    <style>
    * {font-family: 'Courier New', monospace; box-sizing: border-box;}
    TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
    TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #FFD700;}
    TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
    </style>
"@


class Main {
    static [string]$Path = "OneNote x Powershell\Reports\"
    static [string]$Style

    [List[Notebook]]$Notebooks

    Main() {
        # Init date helper class before use
        [DateHelper]::Init()

        $this.Notebooks = [List[Notebook]]::new()

        # Gets all OneNote things
        $onenote = New-Object -ComObject OneNote.Application
        $schema = @{one=”http://schemas.microsoft.com/office/onenote/2013/onenote”}
        [xml]$hierarchy = ""
        $onenote.GetHierarchy("", [OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

        foreach ($notebookXml in $hierarchy.Notebooks.Notebook) {
            # Exclude the admin notebook
            if ($notebookXml.Name.Contains("QuestLearning's")) {
                continue
            }

            $this.Notebooks.Add([Notebook]::new($notebookXml))
        }
    }

    [string]FullReport() {
        [Indenter]$indenter = [Indenter]::new()

        # Prints each notebook in the list
        foreach ($notebook in $this.Notebooks) {
            $indenter += $notebook.FullReport()
        }

        $str = $indenter.Print()
        Set-Content -Path ([Main]::Path + "FULL REPORT.txt") -Value $str
        return $str
    }

    FullReportHtml() {
        # Tough implementation good luck

    }

    [string]StatusReport([Func[Notebook,List[Page]]]$func, [string]$name) {
        [Indenter]$indenter = [Indenter]::new()
        [List[Page]]$pages = [List[Page]]::new()

        foreach ($notebook in $this.Notebooks) {
            [List[Page]]$list = $func.Invoke($notebook)
            if ($list -eq $null) { continue }

            $pages.AddRange($list)
        }

        $indenter += $pages.Count.ToString() + " " + $name
        foreach ($page in $pages) {
            $indenter += $page.ToString()
        }

        return $indenter.Print()
    }

    [string]StatusReports() {
        [Indenter]$indenter = [Indenter]::new()

        $indenter +=
            $this.StatusReport({param([Notebook]$n) $n.GetUngradedPages()},   "ungraded pages"),   " ",
            $this.StatusReport({param([Notebook]$n) $n.GetInactivePages()},   "inactive pages"),   " ",
            $this.StatusReport({param([Notebook]$n) $n.GetEmptyPages()},      "empty pages"),      " ",
            $this.StatusReport({param([Notebook]$n) $n.GetUnreviewedPages()}, "unreviewed pages"), " "

        $str = $indenter.Print()
        Set-Content -Path ([Main]::Path + "STATUS REPORT.txt") -Value $str
        return $str
    }

    StatusReportHtml([Func[Notebook,List[Page]]]$func, [string]$name) {
        [PageHtml[]]$pages = @()

        foreach ($notebook in $this.Notebooks) {
            [List[Page]]$list = $func.Invoke($notebook)
            if ($list -eq $null) { continue }

            foreach ($page in $func.Invoke($notebook)) {
                $pages += [PageHtml]::new($page)
            }
        }

        $html = $pages | ConvertTo-Html -As Table -Head ([HtmlManager]::GetFullHead($name)) -PreContent ([HtmlManager]::MakePre($name))
        Set-Content -Path ("OneNote x Powershell\Reports\" + $name + ".html") -Value $html
    }

    StatusReportsHtml() {
        $this.StatusReportHtml({param([Notebook]$n) $n.GetUngradedPages()},   "UngradedPages")
        $this.StatusReportHtml({param([Notebook]$n) $n.GetInactivePages()},   "InactivePages")
        $this.StatusReportHtml({param([Notebook]$n) $n.GetEmptyPages()},      "EmptyPages")
        $this.StatusReportHtml({param([Notebook]$n) $n.GetUnreviewedPages()}, "UnreviewedPages")
    }

    [string]MissingAssignmentReport() {
        [Indenter]$indenter = [Indenter]::new()

        $sundayskip = 0
        for([int]$i = 0; $i -lt 3; $i++) {
            [datetime]$date = [DateHelper]::Today.AddDays($i + $sundayskip)
            if ([DateHelper]::IsSameWeekday($date, "SUNDAY")) {
                # No assignments on sundays: skip this day
                $sundayskip += 1
                $date = $date.AddDays(1)
            }

            $indenter += ($date.ToString().Substring(0, $date.ToString().IndexOf(" ")) + " missing:")
            $indenter.IncreaseIndent("    - ")

            foreach ($notebook in $this.Notebooks) {
                if (-not $notebook.HasAssignmentOn($date)) {
                    $indenter += $notebook.Name
                }
            }

            $indenter.DecreaseIndent()
            $indenter += " "
        }
        
        $str = $indenter.Print()
        Set-Content -Path ([Main]::Path + "MISSING ASSIGNMENT REPORT.txt") -Value $str
        return $str
    }
    
}



[Main]$main = [Main]::new()
$main.FullReport()
" "
" "
$main.StatusReports()
$main.StatusReportsHtml()
" "
" "
$main.MissingAssignmentReport()