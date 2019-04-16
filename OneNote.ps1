#Require -Version 5.0
using namespace Microsoft.Office.InterOp
using namespace System.Collections.Generic
using namespace System.Xml

# Magic word clears the console
cls



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
        if ($this.Left -gt $r.Right) {
            return $false
        } elseif ($r.Left -gt $this.Right) {
            return $false
        } elseif ($this.Top -gt $r.Bottom) {
            return $false
        } elseif ($r.Top -gt $this.Bottom) {
            return $false
        }
        return $true
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
        $this.Indents = [List[string]]::new()
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
        If ($this.Indents.Count -gt 0) {
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

    ClearLines() {
        $this.Lines = [List[string]]::new()
    }
}


#############
# INK CLASS #
############# (because OneNote is too fancy)
class Ink {
    static [bool]$Debug = $false

    [Rectangle]$Rect
    [string]$Text

    Ink([XmlElement]$ink, [bool]$isWord) {
        If ($isWord) {
            $this.Rect = [Rectangle]::new(-$ink.inkOriginX, -$ink.inkOriginY, $ink.width, $ink.height)
            $this.Text = "[Text]: " + $ink.recognizedText
        } Else {
            $this.Rect = [Rectangle]::new($ink.Position.X, $ink.Position.Y, $ink.Size.Width, $ink.Size.Height)
            $this.Text = "[Drawing]"
        }
    }

    [string]ToString() {
        return $this.Text +
            $(If ($this.Text.Length -gt 0) { " " } Else { "" }) +
            $(If ([Ink]::Debug) { $this.Rect.ToString() } Else { "" })
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

        $imageDisplay = $this.Rect.ToString() # + " " + $this.InkArea + " " + $this.Rect.GetArea() uncomment to evaluate area proportions
        If ($this.HasWork) {
            $imageDisplay += " (!)(has work)"
        }
        $indenter.AddLine($imageDisplay)
            
        if ($this.Inks.Count -gt 0) {
            # INK HEADER
            $indenter.AddLine([string]$this.Inks.Count + " inks:")

            # Ink print
            $inkIndex = 1
            $indenter.IncreaseIndent("|   ")
            foreach ($ink in $this.Inks) {
                $indenter.AddLine([string]$inkIndex + ") " + $ink.ToString())
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

    [datetime]$LastAssignedTime
    [datetime]$LastModifiedTime
    [string]$DateDisplay

    [bool]$Active
    [bool]$Changed
    [bool]$HasWork
    
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
        $this.LastModifiedTime = [datetime]$page.lastModifiedTime
        $this.DateDisplay = $page.lastModifiedTime
        if ($this.TagName -eq [Page]::DefaultTagName) {
            $this.LastAssignedTime = [datetime]$page.dateTime
        } else {
            $this.LastAssignedTime = [datetime]$this.Tag.creationDate
        }

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
                If ($ink.Rect.Intersects($theImage.Rect)) {
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
        $this.Active = ($this.LastModifiedTime -gt (Get-Date).AddDays(-1 * [Page]::ActiveThreshold))
        $this.Changed = ($this.LastModifiedTime -gt $this.LastAssignedTime)
        $this.HasWork = ($this.Images.Where({$_.HasWork -eq $true}).Count -gt 0)
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()

        $statusDisplay = $this.DateDisplay
        If ($this.NeedsGrading -eq $true) {
            $statusDisplay += " (!)(needs grading)"
        }
        ElseIf ($this.Changed -eq $true) {
            $statusDisplay += " (!)(modified)"
        }

        $indenter.AddLine($this.Name.PadRight(40) + " " + $statusDisplay)
        $indenter.IncreaseIndent()

        # Header print
        $indenter.AddLine($this.TagName)
        $indenter.AddLine([string]$this.Images.Count + " image(s):")

        # Image print
        $imageIndex = 1
        $indenter.IncreaseIndent("|   ")
        foreach ($image in $this.Images) {
            $indenter.AddLine([string]$imageIndex + ") " + $image.FullReport())
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
        If ($this.Deleted -eq $true) {
            $sectionDisplay += " (deleted)"
        }
        $indenter.AddLine($sectionDisplay)

        # Page print
        $indenter.IncreaseIndent()
        foreach ($page in $this.Pages) {
            $indenter.AddLine($page.FullReport())
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

    [List[Page]]GetUngradedPages() {
        [List[Page]]$pagesNeedingGrading = [List[Page]]::new()
        foreach ($section in $this.Sections) {
            foreach ($page in $section.Pages.Where({($_.Changed -eq $true) -and ($_.HasWork -eq $true)})) {
                $pagesNeedingGrading.Add($page)
            }
        }
        return $pagesNeedingGrading
    }

    [List[Page]]GetInactivePages() {
        [List[Page]]$pagesInactive = [List[Page]]::new()
        foreach ($section in $this.Sections) {
            foreach ($page in $section.Pages.Where({$_.Active -eq $false})) {
                $pagesInactive.Add($page)
            }
        }
        return $pagesInactive
    }

    [List[Page]]GetEmptyPages() {
        [List[Page]]$pagesEmpty = [List[Page]]::new()
        foreach ($section in $this.Sections) {
            foreach ($page in $section.Pages.Where({$_.Images.Count -eq 0})) {
                $pagesEmpty.Add($page)
            }
        }
        return $pagesEmpty
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()

        $indenter.AddLine(" ")
        $indenter.AddLine($this.Name)
        $indenter.AddLine("-------------------")

        foreach ($section in $this.Sections) {
            $indenter.AddLine($section.FullReport())
        }

        return $indenter.Print()
    }
}



##################
# MAIN TRAVERSAL #
##################
function Main {
    # Gets all OneNote things
    $onenote = New-Object -ComObject OneNote.Application
    $schema = @{one=”http://schemas.microsoft.com/office/onenote/2013/onenote”}
    [xml]$hierarchy = ""
    $onenote.GetHierarchy("", [OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

    [List[Page]]$ungradedPages = [List[Page]]::new()
    [List[Page]]$inactivePages = [List[Page]]::new()
    [List[Page]]$emptyPages = [List[Page]]::new()

    # Traverses each notebook and prints each section
    foreach ($notebookXml in $hierarchy.Notebooks.Notebook) {
        if ($notebookXml.Name.Contains("QuestLearning's")) {
            continue
        }

        [Notebook]$notebook = [Notebook]::new($notebookXml)
        $notebook.FullReport()

        # Get special pages for the full list
        foreach ($page in $notebook.GetUngradedPages()) {
            $ungradedPages.Add($page)
        }
        foreach ($page in $notebook.GetInactivePages()) {
            $inactivePages.Add($page)
        }
        foreach ($page in $notebook.GetEmptyPages()) {
            $emptyPages.Add($page)
        }
    }

    
    " "
    $ungradedPages.Count.ToString() + " ungraded pages"
    foreach ($page in $ungradedPages) {
        $page.ToString()
    }

    " "
    $inactivePages.Count.ToString() + " inactive pages"
    foreach ($page in $inactivePages) {
        $page.ToString()
    }

    " "
    $emptyPages.Count.ToString() + " empty pages"
    foreach ($page in $emptyPages) {
        $page.ToString()
    }
    
}


$str = Main
Set-Content -Path "OneNote x Powershell\FULLREPORT.txt" -Value $str
$str