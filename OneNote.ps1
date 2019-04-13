﻿#Require -Version 5.0
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
    [bool]$HasWork

    Image([XmlElement]$image) {
        $this.Rect = [Rectangle]::new($image.Position.X, $image.Position.Y, $image.Size.Width, $image.Size.Height)
    }

    SetInk([List[Ink]]$theInks) {
        $this.Inks = $theInks

        $this.InkArea = 0;
        $this.Inks.ToArray().ForEach({$this.InkArea += $_.Rect.GetArea()})

        $this.HasWork = $this.InkArea -ge $this.Rect.GetArea() * [Image]::pageFillConstant
    }

    [string]ToString() {
        $lines = [List[string]]::new()
        $indenter = [Indenter]::new()

        $imageDisplay = $this.Rect.ToString() # + " " + $this.InkArea + " " + $this.Rect.GetArea() uncomment to evaluate area proportions
        If ($this.HasWork) {
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
            $this.Tag = $tagDefs[0]
            $this.TagName = $this.Tag.Name
        }
        else {
            $this.TagName = [Page]::DefaultTagName
        }

        # Get dates
        $this.LastModifiedTime = [datetime]$page.lastModifiedTime
        $this.DateDisplay = $page.lastModifiedTime
        if ($this.TagName -eq [Page]::DefaultTagName) {
            $this.LastAssignedTime = [datetime]$this.Tag.creationDate
        } else {
            $this.LastAssignedTime = [datetime]$page.dateTime
        }

        # Finds main page content
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

        # Determine the status of the page
        $this.Active = ($this.LastModifiedTime -gt (Get-Date).AddDays(-1 * [Page]::ActiveThreshold))
        $this.Changed = ($this.LastModifiedTime -gt $this.LastAssignedTime)
        $this.HasWork = ($this.Images.Where({$_.HasWork -eq $true}).Count -gt 0)
    }

    [string]ToString() {
        $lines = [List[string]]::new()
        $indenter = [Indenter]::new()

        $statusDisplay = $this.DateDisplay
        If ($this.NeedsGrading -eq $true) {
            $statusDisplay += " (!)(needs grading)"
        }
        ElseIf ($this.Changed -eq $true) {
            $statusDisplay += " (!)(modified)"
        }

        $lines.Add($this.Name.PadRight(40) + " " + $statusDisplay)
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
            $onenote.GetPageContent($pageXml.ID, [ref]$content, [OneNote.PageInfo]::piBasic)

            $this.Pages.Add([Page]::new($pageXml, $content, $this))
        }
    }

    [List[Page]]PagesNeedingGrading() {
        [List[Page]]$pagesNeedingGrading = [List[Page]]::new()
        $this.Pages.Where({($_.Changed -eq $true) -and ($_.HasWork -eq $true)}).ForEach({$pagesNeedingGrading.Add($_)})
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

        return $lines -join "`r`n"
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

    [List[Page]]$pagesNeedingGrading = [List[Page]]::new()

    # Traverses each notebook and prints each section
    foreach ($notebook in $hierarchy.Notebooks.Notebook) {
        " "
        $notebook.Name
        "-----------------"

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

        # Goes through each section and obtains / prints relevant information
        foreach ($sectionXml in $sectionXmls) {
            [Section]$section = [Section]::new($sectionXml, $notebook.Name)
            $section.ToString()

            [List[Page]]$thePagesNeedingGrading = $section.PagesNeedingGrading();
            foreach ($pageNeedingGrading in $thePagesNeedingGrading) {
                $pagesNeedingGrading.Add($pageNeedingGrading)
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