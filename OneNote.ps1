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

        $imageDisplay = $this.Rect.ToString() # + " " + $this.InkArea + " " + $this.Rect.GetArea() uncomment to evaluate area proportions
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
        $this.Active = ($this.LastModifiedTime -gt (Get-Date).AddDays(-1 * [Page]::ActiveThreshold))
        $this.Changed = ($this.LastModifiedTime -gt $this.LastAssignedTime)
        $this.HasWork = ($this.Images.Where({$_.HasWork -eq $true}).Count -gt 0)
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

    [List[Page]]GetUngradedPages() {
        return $this.GetPagesWhere({param([Page]$p) ($p.Changed -eq $true) -and ($p.HasWork -eq $true)})
    }
    [List[Page]]GetInactivePages() {
        return $this.GetPagesWhere({param([Page]$p) $p.Active -eq $false})
    }
    [List[Page]]GetEmptyPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.Images.Count -eq 0})
    }
    [List[Page]]GetUnreviewedPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.TagName -like "*REVIEW*"})
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
        Set-Content -Path "OneNote x Powershell\FULLREPORT.txt" -Value $str
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

        $indenter += " ", ($pages.Count.ToString() + " " + $name)
        foreach ($page in $pages) {
            $indenter += $page.ToString()
        }

        return $indenter.Print()
    }

    [string]StatusReports() {
        [Indenter]$indenter = [Indenter]::new()

        $indenter += $this.StatusReport({param([Notebook]$n) $n.GetUngradedPages()},   "ungraded pages"),
            $this.StatusReport({param([Notebook]$n) $n.GetInactivePages()},   "inactive pages"),
            $this.StatusReport({param([Notebook]$n) $n.GetEmptyPages()},      "empty pages"),
            $this.StatusReport({param([Notebook]$n) $n.GetUnreviewedPages()}, "unreviewed pages")

        $str = $indenter.Print()
        Set-Content -Path "OneNote x Powershell\STATUSREPORT.txt" -Value $str
        return $str
    }
    
}


[Main]$main = [Main]::new()
$main.FullReport()
" "
$main.StatusReports()