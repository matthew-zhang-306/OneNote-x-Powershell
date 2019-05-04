#Require -Version 5.0
using namespace Microsoft.Office.Interop
using namespace System.Collections.Generic
using namespace System.Xml

# Magic word clears the console
cls


# If Powershell for some reason doesn't recognize OneNote classes, type this into the console to magically fix it
# Add-Type -AssemblyName Microsoft.Office.Interop.OneNote



###################
# RECTANGLE CLASS #
###################
class Rectangle {
    [float]$X
    [float]$Y
    [float]$Width
    [float]$Height

    [float]$Left
    [float]$Right
    [float]$Top
    [float]$Bottom

    Rectangle([float]$x, [float]$y, [float]$w, [float]$h) {
        $this.X = $x
        $this.Y = $y
        $this.Width = $w
        $this.Height = $h

        $this.Left = $x
        $this.Right = $x + $w
        $this.Top = $y
        $this.Bottom = $y + $h
    }

    [bool]Intersects([Rectangle]$r) {
        return -not (($this.Left -gt $r.Right) -or ($r.Left -gt $this.Right) -or ($this.Top -gt $r.Bottom) -or ($r.Top -gt $this.Bottom))
    }

    [float]GetArea() {
        return $this.Width * $this.Height
    }

    [string]ToString() {
        return "RECT[" + $this.X + ", " + $this.Y + ", " + $this.Width + ", " + $this.Height + "]"
    }
}



####################
# INDENTER UTILITY #
####################
class Indenter {
    [List[string]]$Indents
    [List[string]]$Lines

    Indenter() {
        $this.ClearLines()
    }

    [string]Print() {
        return $this.Print($this.Lines)
    }
    [string]Print([string]$outputRaw) {
        $output = [List[string]]::new()
        foreach ($line in ($outputRaw -split '\r?\n')) {
            $output.Add($this.GetCurrentIndent() + $line)
        }
        return $this.Print($output)
    }
    [string]Print([List[string]]$output) {
        # note to self: when splitting strings only '\r\n' works, but when joining strings only "`r`n" works. the inconsistency is weird
        return $output -join "`r`n"
    }

    IncreaseIndent() {
        $this.IncreaseIndent("    ")
    }
    IncreaseIndent([string]$indent) {
        $this.Indents.Add($indent)
    }

    DecreaseIndent() {
        if ($this.Indents.Count -gt 0) {
            $this.Indents.RemoveAt($this.Indents.Count - 1)
        }
    }

    [string]GetCurrentIndent() {
        return $this.Indents -join ""
    }

    AddLine([string]$line) {
        $this.AddLines($line -split '\r?\n')
    }
    AddLines([List[string]]$lines) {
        foreach($line in $lines) {
            $this.Lines.Add($this.GetCurrentIndent() + $line)
        }
    }

    # Note: Do not use + (since this function is mutating), instead use +=
    static [Indenter]op_Addition([Indenter]$first, [string]$second) {
        $first.AddLine($second.ToString())
        return $first
    }
    static [Indenter]op_Addition([Indenter]$first, [System.Array]$second) {
        foreach ($item in $second) {
            $first += $item
        }
        return $first
    }
    
    ClearLines() {
        $this.Indents = [List[string]]::new()
        $this.Lines = [List[string]]::new()
    }
}



####################
# DATEHELPER CLASS #
####################
class DateHelper {
    static [Dictionary[string, int]]$WeekdayMap
    static Init() {
        [DateHelper]::WeekdayMap = [Dictionary[string, int]]::new()
        [DateHelper]::WeekdayMap.Add("Monday", 1)
        [DateHelper]::WeekdayMap.Add("Tuesday", 2)
        [DateHelper]::WeekdayMap.Add("Wednesday", 3)
        [DateHelper]::WeekdayMap.Add("Thursday", 4)
        [DateHelper]::WeekdayMap.Add("Friday", 5)
        [DateHelper]::WeekdayMap.Add("Saturday", 6)
        [DateHelper]::WeekdayMap.Add("Sunday", 7)
    }
    
    static [datetime]$Now = (Get-Date -Year 2019 -Month 4 -Day 1) # Comment out parameters to use the current date and not a debug time
    static [datetime]$Today = [DateHelper]::Now.Date

    static [bool]IsSameDay([datetime]$date1, [datetime]$date2) {
        return $date1.Date.ToString() -eq $date2.Date.ToString()
    }

    static [bool]IsValidWeekday([string]$weekday) {
        return [DateHelper]::WeekdayMap.ContainsKey([DateHelper]::PascalCase($weekday))
    }
    static [string]GetWeekday([datetime]$date) {
        return $date.DayOfWeek.ToString()
    }
    static [bool]IsSameWeekday([datetime]$date1, [datetime]$date2) {
        return [DateHelper]::GetWeekday($date1) -eq [DateHelper]::GetWeekday($date2)
    }
    static [bool]IsSameWeekday([datetime]$date, [string]$dateStr) {
        return [DateHelper]::GetWeekday($date) -eq [DateHelper]::PascalCase($dateStr)
    }

    # Meant to convert raw weekday strings into formalized ones (eg "MONDAY" => "Monday") for comparison
    static [string]PascalCase([string]$str) {
        return $str.Substring(0, 1).ToUpper() + $str.Substring(1).ToLower()
    }
}
# Init date helper class before use
[DateHelper]::Init()



#############
# INK CLASS #
############# (because OneNote is too fancy)
class Ink {
    static [bool]$Debug = $false

    [Rectangle]$Rect
    [string]$Text

    Ink([XmlElement]$ink, [bool]$isWord) {
        if ($isWord) {
            $this.Rect = [Rectangle]::new(-$ink.inkOriginX, -$ink.inkOriginY, $ink.width, $ink.height)
            $this.Text = "[Text]: " + $ink.recognizedText
        } else {
            $this.Rect = [Rectangle]::new($ink.Position.X, $ink.Position.Y, $ink.Size.Width, $ink.Size.Height)
            $this.Text = "[Drawing]"
        }
    }

    [string]ToString() {
        return $this.Text +
            $(if ($this.Text.Length -gt 0) { " " } else { "" }) +
            $(if ([Ink]::Debug) { $this.Rect.ToString() } else { "" })
    }
}



###############
# IMAGE CLASS #
###############
class Image {
    static [float] $pageFillConstant = 0.005
    
    [Rectangle]$Rect
    [List[Ink]]$Inks
    [float]$InkArea
    [bool]$HasWork

    Image([XmlElement]$image) {
        $this.Rect = [Rectangle]::new($image.Position.X, $image.Position.Y, $image.Size.Width, $image.Size.Height)
    }

    SetInk([List[Ink]]$theInks) {
        $this.Inks = $theInks

        $this.InkArea = 0;
        foreach ($ink in $this.Inks.ToArray()) {
            $this.InkArea += $ink.Rect.GetArea()
        }

        $this.HasWork = $this.InkArea -ge $this.Rect.GetArea() * [Image]::pageFillConstant
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()

        $imageDisplay = $this.Rect.ToString() # + " " + $this.InkArea + " " + $this.Rect.GetArea() uncomment to debug area proportions
        if ($this.HasWork) {
            $imageDisplay += " (!)(has work)"
        }
        $indenter += $imageDisplay
            
        if ($this.Inks.Count -gt 0) {
            # INK HEADER
            $indenter += [string]$this.Inks.Count + " inks:"

            # Ink print
            $inkIndex = 1
            $indenter.IncreaseIndent("|   ")
            foreach ($ink in $this.Inks) {
                $indenter += [string]$inkIndex + ") " + $ink.ToString()
                $inkIndex += 1
            }
            $indenter.DecreaseIndent()
        }
        
        return $indenter.Print()
    }
}



##############
# PAGE CLASS #
##############
class Page {
    static [float]$ActiveThreshold = 3.0
    static [string]$DefaultTagName = "# NoTag #"

    [string]$Name
    [string]$TagName
    [XmlElement]$Tag

    [datetime]$CreationTime
    [datetime]$LastAssignedTime
    [datetime]$LastModifiedTime
    [string]$DateDisplay

    [datetime]$OriginalAssignmentDate

    [bool]$Active
    [bool]$Changed
    [bool]$HasWork
    [bool]$Empty
    
    [List[Image]]$Images
    [List[Ink]]$Inks
    [Section]$Section

    Page([XmlElement]$page, [xml]$content, [Section]$section) {
        $this.Name = $page.Name
        $this.Section = $section

        # Get tag
        [XmlElement[]]$tags = $content.GetElementsByTagName("one:Tag")
        [XmlElement[]]$tagDefs = $content.GetElementsByTagName("one:TagDef")
        if (($tags.Length -gt 0) -and ($tagDefs.Length -gt 0)) {
            $this.Tag = $tags[0]
            $this.TagName = $tagDefs[0].Name
        }
        else {
            $this.TagName = [Page]::DefaultTagName
        }

        # Get dates
        $this.CreationTime = [datetime]$page.dateTime
        $this.LastModifiedTime = [datetime]$page.lastModifiedTime
        if ($this.TagName -eq [Page]::DefaultTagName) {
            $this.LastAssignedTime = [datetime]$this.CreationTime
        } else {
            $this.LastAssignedTime = [datetime]$this.Tag.creationDate
        }

        if ([DateHelper]::IsValidWeekday($this.Section.Name)) {
            $this.OriginalAssignmentDate = $this.CreationTime.Date
            while (-not [DateHelper]::IsSameWeekday($this.OriginalAssignmentDate, $this.Section.Name)) {
                $this.OriginalAssignmentDate = $this.OriginalAssignmentDate.AddDays(1)
            }
        }
        else {
            $this.OriginalAssignmentDate = $this.LastAssignedTime
        }

        $this.DateDisplay = $this.OriginalAssignmentDate

        # Finds main page content
        $this.Inks = [List[Ink]]::new()
        foreach ($ink in $content.GetElementsByTagName("one:InkDrawing")) {
            $this.Inks.Add([Ink]::new($ink, $false))
        }
        foreach ($ink in $content.GetElementsByTagName("one:InkWord")) {
            $this.Inks.Add([Ink]::new($ink, $true))
        }

        $this.Images = [List[Image]]::new()
        foreach ($image in $content.GetElementsByTagName("one:Image").Where{!($_.Position -eq $null)}) {
            $theImage = [Image]::new($image)

            # Get contained inks
            $theInks = [List[Ink]]::new()
            foreach ($ink in $this.Inks.ToArray()) {
                if ($ink.Rect.Intersects($theImage.Rect)) {
                    $theInks.Add($ink)
                }
            }
            $theImage.SetInk($theInks)
            
            $this.Images.Add($theImage)
        }

        # Debug log full XML
        if ($page.name.StartsWith("Quest2-B_answerkey")) { # <-- change this string
            Set-Content -Path "OneNote x Powershell\log.txt" -Value $content.InnerXml
        }

        # Determine the status of the page
        $this.Active = ($this.LastModifiedTime -gt [DateHelper]::Now.AddDays(-1 * [Page]::ActiveThreshold))
        $this.Changed = ($this.LastModifiedTime -gt $this.LastAssignedTime)
        $this.HasWork = ($this.Images.Where({$_.HasWork -eq $true}).Count -gt 0)
        $this.Empty = ($this.Images.Count -eq 0)
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()

        $statusDisplay = $this.DateDisplay
        if ($this.NeedsGrading -eq $true) {
            $statusDisplay += " (!)(needs grading)"
        }
        elseif ($this.Changed -eq $true) {
            $statusDisplay += " (!)(modified)"
        }

        $indenter += $this.Name.PadRight(40) + " " + $statusDisplay
        $indenter.IncreaseIndent()

        # Header print
        $indenter += $this.TagName, ([string]$this.Images.Count + " image(s):")

        # Image print
        $imageIndex = 1
        $indenter.IncreaseIndent("|   ")
        foreach ($image in $this.Images) {
            $indenter += [string]$imageIndex + ") " + $image.FullReport()
            $imageIndex += 1
        }
        $indenter.DecreaseIndent()

        $indenter.DecreaseIndent()
        return $indenter.Print()
    }

    [string]ToString() {
        return "PAGE: " + $this.Section.Notebook.Name.PadRight(40) + " | " + $this.Section.Name.PadRight(40) + " | " + $this.Name
    }
}



#################
# SECTION CLASS #
#################
class Section {
    [string]$Name
    [bool]$Deleted
    [List[Page]]$Pages
    [Notebook]$Notebook

    Section([XmlElement]$section, [Notebook]$notebook) {
        $this.Name = $section.Name
        $this.Deleted = $section.IsInRecycleBin
        $this.Notebook = $notebook

        $this.Pages = [List[Page]]::new()
        foreach ($pageXml in $section.Page) {
            # We cannot pass a ComObject as a parameter and still have it work, so it is redefined here
            $onenote = New-Object -ComObject OneNote.Application

            # Get page content
            [xml]$content = ""
            $onenote.GetPageContent($pageXml.ID, [ref]$content, [OneNote.PageInfo]::piBasic)

            $this.Pages.Add([Page]::new($pageXml, $content, $this))
        }
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()
        
        # Header print
        $sectionDisplay = "# Section: " + $this.Name + " #"
        if ($this.Deleted -eq $true) {
            $sectionDisplay += " (deleted)"
        }
        $indenter += $sectionDisplay

        # Page print
        $indenter.IncreaseIndent()
        foreach ($page in $this.Pages) {
            $indenter += $page.FullReport()
        }
        $indenter.DecreaseIndent()

        return $indenter.Print()
    }
}



##################
# NOTEBOOK CLASS #
##################
class Notebook {
    [string]$Name
    [bool]$Deleted
    [List[Section]]$Sections

    Notebook([XmlElement]$notebook) {
        $this.Name = $notebook.Name
        $this.Deleted = $notebook.IsInRecycleBin

        [List[XmlElement]]$sectionXmls = [List[XmlElement]]::new()
        # Checks for all sections placed in a sectiongroup
        foreach ($sectiongroup in $notebook.SectionGroup) {
            if ($sectiongroup.isInRecycleBin -eq $false) {
                foreach ($sectionXml in $sectiongroup.Section) {
                    $sectionXmls.Add($sectionXml)
                }
            }
        }
        # Checks for any sections not placed in a sectiongroup
        foreach ($sectionXml in $notebook.Section) {
            $sectionXmls.Add($sectionXml)
        }

        # Goes through all the xml pieces and makes section objects
        $this.Sections = [List[Section]]::new()
        foreach ($sectionXml in $sectionXmls) {
            $this.Sections.Add([Section]::new($sectionXml, $this))
        }
    }

    [List[Page]]GetPagesWhere([Func[Page,bool]]$func) {
        [List[Page]]$pages = [List[Page]]::new()
        foreach ($section in $this.Sections) {
            foreach ($page in $section.Pages) {
                if ($func.Invoke($page)) {
                    $pages.Add($page)
                }
            }
        }
        return $pages
    }
    [bool]HasPagesWhere([Func[Page,bool]]$func) {
        foreach ($section in $this.Sections) {
            foreach ($page in $section.Pages) {
                if ($func.Invoke($page)) {
                    return $true
                }
            }
        }
        return $false
    }

    [List[Page]]GetUngradedPages() {
        return $this.GetPagesWhere({param([Page]$p) ($p.Changed -eq $true) -and ($p.HasWork -eq $true)})
    }
    [List[Page]]GetInactivePages() {
        return $this.GetPagesWhere({param([Page]$p) $p.Active -eq $false})
    }
    [List[Page]]GetEmptyPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.Empty -eq $true})
    }
    [List[Page]]GetUnreviewedPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.TagName -like "*REVIEW*"})
    }

    [bool]HasAssignmentOn([datetime]$date) {
        return $this.HasPagesWhere({param([Page]$p) ([DateHelper]::IsSameDay($p.OriginalAssignmentDate, $date)) -and (-not $p.Empty)})
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()

        $indenter += " ", $this.Name, "-------------------"
        foreach ($section in $this.Sections) {
            $indenter += $section.FullReport()
        }

        return $indenter.Print()
    }
}



##############
# MAIN CLASS #
##############
class Main {
    static [string]$Path = "OneNote x Powershell\Reports\"
    [List[Notebook]]$Notebooks

    Main() {
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

    [string]StatusReport([Func[Notebook,List[Page]]]$func, [string]$name) {
        [Indenter]$indenter = [Indenter]::new()
        [List[Page]]$pages = [List[Page]]::new()

        foreach ($notebook in $this.Notebooks) {
            [List[Page]]$list = $func.Invoke($notebook)
            if ($list -eq $null) {
                # Debug here because to be honest it really should not be null
            } else {
                $pages.AddRange($func.Invoke($notebook))
            }
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
" "
" "
$main.MissingAssignmentReport()