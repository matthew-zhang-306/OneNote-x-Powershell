#Require -Version 5.0
using namespace Microsoft.Office.InterOp
using namespace System.Collections.Generic
using namespace System.Xml

# Magic word clears the console
cls

$numDays = 1

# Gets all one note things
$onenote = New-Object -ComObject OneNote.Application
$schema = @{one=”http://schemas.microsoft.com/office/onenote/2013/onenote”}
[xml]$hierarchy = ""
$onenote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)



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

    Indenter() {
        $this.Indents = [List[string]]::new()
    }

    [string]Print([string]$output) {
        $lines = [List[string]]::new()
        ($output -split '\r?\n').ForEach({$lines.Add(($this.Indents -join "") + $_)})
        return $lines -join "`r`n"
    }                                    # note to self: when splitting strings only '\r\n' works, but when joining strings only "`r`n" works. the inconsistency is weird

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
}


#############
# INK CLASS #
############# (because OneNote is too fancy)
class Ink {
    [Rectangle]$Rect
    [string]$Text

    Ink([XmlElement]$ink, [bool]$isWord) {
        If ($isWord) {
            $this.Rect = [Rectangle]::new(-$ink.inkOriginX, -$ink.inkOriginY, $ink.width, $ink.height)
            $this.Text = $ink.recognizedText
        } Else {
            $this.Rect = [Rectangle]::new($ink.Position.X, $ink.Position.Y, $ink.Size.Width, $ink.Size.Height)
            $this.Text = ""
        }
    }

    [string]ToString() {
        return $this.Text + $(If ($this.Text.Length -gt 0) { " " } Else { "" }) + $this.Rect.ToString()
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

    Image([XmlElement]$image) {
        $this.Rect = [Rectangle]::new($image.Position.X, $image.Position.Y, $image.Size.Width, $image.Size.Height)
    }

    SetInk([List[Ink]]$theInks) {
        $this.Inks = $theInks

        $this.InkArea = 0;
        $this.Inks.ToArray().ForEach({$this.InkArea += $_.Rect.GetArea()})
    }

    [string]ToString() {
        $lines = [List[string]]::new()
        $indenter = [Indenter]::new()

        $imageDisplay = $this.Rect.ToString() # + " " + $this.InkArea + " " + $this.Rect.GetArea() uncomment to evaluate area proportions
        If ($this.InkArea -ge $this.Rect.GetArea() * [Image]::pageFillConstant) {
            $imageDisplay += " (!)(has work)"
        }
        $lines.Add($imageDisplay)
            
        if ($this.Inks.Count -gt 0) {
            # INK HEADER
            $lines.Add([string]$this.Inks.Count + " inks:")

            # Ink print
            $inkIndex = 1
            $indenter.IncreaseIndent("|   ")
            foreach ($ink in $this.Inks) {
                $lines.Add($indenter.Print([string]$inkIndex + ") " + $ink.ToString()))
                $inkIndex += 1
            }
            $indenter.DecreaseIndent()
        }
        
        return $lines -join "`r`n"
    }
}


##############
# PAGE CLASS #
##############
class Page {
    static [int] $dateModifiedThreshold = 1

    [string]$Name
    [string]$Tag
    [string]$DateDisplay
    [bool]$Changed
<<<<<<< HEAD
    [bool]$NeedsGrading
    [List[Image]]$Images
    [List[Ink]]$Inks
    [Section]$Section
=======
    [System.Collections.Generic.List[Image]]$Images
    [System.Collections.Generic.List[Ink]]$Inks
>>>>>>> parent of 141ebd9... Section Class

    Page([XmlElement]$page, [xml]$content, [Section]$section) {
        $this.Name = $page.Name
        $this.Section = $section

        # Determine if the last modified date is recent enough
        $this.DateDisplay = $page.lastModifiedTime
        $this.Changed = $false
        if ([datetime]$page.lastModifiedTime -gt (Get-Date).AddDays(-1 * [Page]::dateModifiedThreshold)) {
            $this.DateDisplay += " (!)(changed)"
            $this.Changed = $true
        }

        # Finds content
        [XmlElement[]]$tags = $content.GetElementsByTagName("one:Tag")
        [XmlElement[]]$tagDefs = $content.GetElementsByTagName("one:TagDef")
        if (($tags.Length -gt 0) -and ($tagDefs.Length -gt 0)) {
            $this.Tag = $tagDefs[0].Name
        }
        else {
            $this.Tag = "No tag"
        }

        $this.Inks = [List[Ink]]::new()
        $content.GetElementsByTagName("one:InkDrawing").ForEach({$this.Inks.Add([Ink]::new($_, $false))})
        $content.GetElementsByTagName("one:InkWord").ForEach({$this.Inks.Add([Ink]::new($_, $true))})

        $this.Images = [List[Image]]::new()
        $content.GetElementsByTagName("one:Image").Where{!($_.Position -eq $null)}.ForEach({
            $theImage = [Image]::new($_)

            # Get contained inks
            $theInks = [List[Ink]]::new()
            $this.Inks.ToArray().ForEach({If ($_.Rect.Intersects($theImage.Rect)) { $theInks.Add($_) }})
            $theImage.SetInk($theInks)
            
            $this.Images.Add($theImage)
        })

        # Debug log full XML
        if ($page.name.StartsWith("Quest2-B_answerkey")) { # <-- change this string
            Set-Content -Path "OneNote x Powershell\log.txt" -Value $content.InnerXml
        }
    }

    [string]ToString() {
        $lines = [List[string]]::new()
        $indenter = [Indenter]::new()

        $lines.Add($this.Name.PadRight(40) + " " + $this.DateDisplay)
        $indenter.IncreaseIndent()

        # Header print
        $lines.Add($indenter.Print($this.Tag))
        $lines.Add($indenter.Print([string]$this.Images.Count + " image(s):"))

        # Image print
        $imageIndex = 1
        $indenter.IncreaseIndent("|   ")
        foreach ($image in $this.Images) {
            $lines.Add($indenter.Print([string]$imageIndex + ") " + $image.ToString()))
            $imageIndex += 1
        }
        $indenter.DecreaseIndent()

        $indenter.DecreaseIndent()
        return $lines -join "`r`n"
    }
}


<<<<<<< HEAD
#################
# SECTION CLASS #
#################
class Section {
    [string]$Name
    [bool]$Deleted
    [List[Page]]$Pages
    [string]$NotebookName

    Section([XmlElement]$section, [string]$notebook) {
        $this.Name = $section.Name
        $this.Deleted = $section.IsInRecycleBin
        $this.NotebookName = $notebook

        $this.Pages = [List[Page]]::new()
        foreach ($pageXml in $section.Page) {
            # We cannot pass a ComObject as a parameter and still have it work, so it is redefined here
            $onenote = New-Object -ComObject OneNote.Application

            # Get page content
            [xml]$content = ""
            $onenote.GetPageContent($pageXml.ID, [ref]$content, [Microsoft.Office.InterOp.OneNote.PageInfo]::piBasic)

            $this.Pages.Add([Page]::new($pageXml, $content, $this))
        }
    }

    [List[Page]]PagesNeedingGrading() {
        [List[Page]]$pagesNeedingGrading = [List[Page]]::new()
        $this.Pages.Where({$_.NeedsGrading -eq $true}).ForEach({$pagesNeedingGrading.Add($_)})
        if ($pagesNeedingGrading.Count -gt 0) {
            
        }
        return $pagesNeedingGrading
    }

    [string]ToString() {
        $lines = [List[string]]::new()
        $indenter = [Indenter]::new()
        
        # Header print
        $sectionDisplay = "# Section: " + $this.Name + " #"
        If ($this.Deleted -eq $true) {
            $sectionDisplay += " (deleted)"
        }
        $lines.Add($sectionDisplay)

        # Page print
        $indenter.IncreaseIndent()
        foreach ($page in $this.Pages) {
            $lines.Add($indenter.Print($page.ToString()))
        }
        $indenter.DecreaseIndent()
=======
##########################
# PRINT SECTION FUNCTION #
##########################
function Print-Section {
    param([System.Xml.XmlElement]$section)

    [Indenter]$indenter = [Indenter]::new()
    $indenter.IncreaseIndent()
    
    # SECTION HEADER
    if ($section.isInRecycleBin -eq $true) {
        $indenter.Print("# Section: " + $section.name + " # (deleted)")
    } else {
        $indenter.Print($indent + "# Section: " + $section.name + " #")
    }

    # PAGE
    $indenter.IncreaseIndent()
    foreach ($pageXml in $section.Page) {
        # Finds important content
        [xml]$content = ""
        $onenote.GetPageContent($pageXml.ID, [ref]$content, [Microsoft.Office.InterOp.OneNote.PageInfo]::piBasic)
>>>>>>> parent of 141ebd9... Section Class

        [Page]$page = [Page]::new($pageXml, $content)
        $indenter.Print($page.ToString())
    }
    $indenter.DecreaseIndent()
}


##################
# MAIN TRAVERSAL #
##################
function Main {
<<<<<<< HEAD
    # Gets all OneNote things
    $onenote = New-Object -ComObject OneNote.Application
    $schema = @{one=”http://schemas.microsoft.com/office/onenote/2013/onenote”}
    [xml]$hierarchy = ""
    $onenote.GetHierarchy("", [OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

    [List[Page]]$pagesNeedingGrading = [List[Page]]::new()

    # Traverses each notebook and prints each section
=======
>>>>>>> parent of 141ebd9... Section Class
    foreach ($notebook in $hierarchy.Notebooks.Notebook) {
        " "
        $notebook.Name
        "-----------------"

        [List[XmlElement]]$sectionXmls = [List[XmlElement]]::new()
        # Checks for all sections placed in a sectiongroup
        foreach ($sectiongroup in $notebook.SectionGroup) {
            if ($sectiongroup.isInRecycleBin -eq $false) {
<<<<<<< HEAD
                foreach ($sectionXml in $sectiongroup.Section) {
                    $sectionXmls.Add($sectionXml)
=======
                "### Section Group: " + $sectiongroup.Name + " ###"
            
                foreach ($section in $sectiongroup.Section) {
                    Print-Section $section
>>>>>>> parent of 141ebd9... Section Class
                }
            }
        }
        # Checks for any sections not placed in a sectiongroup
<<<<<<< HEAD
        foreach ($sectionXml in $notebook.Section) {
            $sectionXmls.Add($sectionXml)
        }

        # Goes through each section and obtains / prints relevant information
        foreach ($sectionXml in $sectionXmls) {
            [Section]$section = [Section]::new($sectionXml, $notebook.Name)
            $section.ToString()

            [List[Page]]$thePagesNeedingGrading = $section.PagesNeedingGrading();
            foreach ($pageNeedingGrading in $thePagesNeedingGrading) {
                $pagesNeedingGrading.Add($pageNeedingGrading)
=======
        $hasMisc = $false
        foreach ($section in $notebook.Section) {
            $hasMisc = $hasMisc -or !($section.isInRecycleBin)
        }
        if ($hasMisc = $true) {
            "### Section Group: Miscellaneous ###"
            foreach ($section in $notebook.Section) {
                Print-Section $section
>>>>>>> parent of 141ebd9... Section Class
            }
        }
    }

    " "
    $pagesNeedingGrading.Count.ToString() + " need grading"
    foreach ($page in $pagesNeedingGrading) {
        "PAGE: " + $page.Section.NotebookName + " " + $page.Section.Name + " " + $page.Name
    }
}

$str = Main
Set-Content -Path "OneNote x Powershell\FULLREPORT.txt" -Value $str
$str