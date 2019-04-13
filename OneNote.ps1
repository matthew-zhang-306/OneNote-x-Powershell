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
    [System.Collections.Generic.List[string]]$Indents

    Indenter() {
        $this.Indents = [System.Collections.Generic.List[string]]::new()
    }

    [string]Print([string]$output) {
        $lines = [System.Collections.Generic.List[string]]::new()
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

    Ink([System.Xml.XmlElement]$ink, [bool]$isWord) {
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
    [System.Collections.Generic.List[Ink]]$Inks
    [float]$InkArea
    [bool]$HasWork

    Image([System.Xml.XmlElement]$image) {
        $this.Rect = [Rectangle]::new($image.Position.X, $image.Position.Y, $image.Size.Width, $image.Size.Height)
    }

    SetInk([System.Collections.Generic.List[Ink]]$theInks) {
        $this.Inks = $theInks

        $this.InkArea = 0;
        $this.Inks.ToArray().ForEach({$this.InkArea += $_.Rect.GetArea()})

        $this.HasWork = $this.InkArea -ge $this.Rect.GetArea() * [Image]::pageFillConstant
    }

    [string]ToString() {
        $lines = [System.Collections.Generic.List[string]]::new()
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
    static [int] $dateModifiedThreshold = 2

    [string]$Name
    [string]$Tag
    [string]$DateDisplay
    [bool]$Changed
    [bool]$NeedsGrading
    [System.Collections.Generic.List[Image]]$Images
    [System.Collections.Generic.List[Ink]]$Inks

    Page([System.Xml.XmlElement]$page, [xml]$content) {
        $this.Name = $page.Name

        # Determine if the last modified date is recent enough
        $this.DateDisplay = $page.lastModifiedTime
        $this.Changed = $false
        if ([datetime]$page.lastModifiedTime -gt (Get-Date).AddDays(-1 * [Page]::dateModifiedThreshold)) {
            $this.Changed = $true
        }

        # Finds content
        [System.Xml.XmlElement[]]$tags = $content.GetElementsByTagName("one:Tag")
        [System.Xml.XmlElement[]]$tagDefs = $content.GetElementsByTagName("one:TagDef")
        if (($tags.Length -gt 0) -and ($tagDefs.Length -gt 0)) {
            $this.Tag = $tagDefs[0].Name
        }
        else {
            $this.Tag = "No tag"
        }

        $this.Inks = [System.Collections.Generic.List[Ink]]::new()
        $content.GetElementsByTagName("one:InkDrawing").ForEach({$this.Inks.Add([Ink]::new($_, $false))})
        $content.GetElementsByTagName("one:InkWord").ForEach({$this.Inks.Add([Ink]::new($_, $true))})

        $this.Images = [System.Collections.Generic.List[Image]]::new()
        $content.GetElementsByTagName("one:Image").Where{!($_.Position -eq $null)}.ForEach({
            $theImage = [Image]::new($_)

            # Get contained inks
            $theInks = [System.Collections.Generic.List[Ink]]::new()
            $this.Inks.ToArray().ForEach({If ($_.Rect.Intersects($theImage.Rect)) { $theInks.Add($_) }})
            $theImage.SetInk($theInks)
            
            $this.Images.Add($theImage)
        })

        # Debug log full XML
        if ($page.name.StartsWith("Quest2-B_answerkey")) { # <-- change this string
            Set-Content -Path "OneNote x Powershell\log.txt" -Value $content.InnerXml
        }

        # Determine if the page has new work
        $this.NeedsGrading = 
            ($this.Tag -eq "No tag") -and
            ($this.Changed -eq $true) -and
            ($this.Images.Where({$_.HasWork -eq $true}).Count -gt 0)
    }

    [string]ToString() {
        $lines = [System.Collections.Generic.List[string]]::new()
        $indenter = [Indenter]::new()

        $statusDisplay = $this.DateDisplay
        If ($this.NeedsGrading -eq $true) {
            $statusDisplay += " (!)(needs grading)"
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
    [System.Collections.Generic.List[Page]]$Pages

    Section([System.Xml.XmlElement]$section) {
        $this.Name = $section.Name
        $this.Deleted = $section.IsInRecycleBin

        $this.Pages = [System.Collections.Generic.List[Page]]::new()
        foreach ($pageXml in $section.Page) {
            # We cannot pass a ComObject as a parameter and still have it work, so it is redefined here
            $onenote = New-Object -ComObject OneNote.Application

            # Get page content
            [xml]$content = ""
            $onenote.GetPageContent($pageXml.ID, [ref]$content, [Microsoft.Office.InterOp.OneNote.PageInfo]::piBasic)

            $this.Pages.Add([Page]::new($pageXml, $content))
        }
    }

    [string]ToString() {
        $lines = [System.Collections.Generic.List[string]]::new()
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
    $onenote.GetHierarchy("", [Microsoft.Office.InterOp.OneNote.HierarchyScope]::hsPages, [ref]$hierarchy)

    # Traverses each notebook and prints each section
    foreach ($notebook in $hierarchy.Notebooks.Notebook) {
        " "
        $notebook.Name
        "-----------------"

        foreach ($sectiongroup in $notebook.SectionGroup) {
            if ($sectiongroup.isInRecycleBin -eq $false) {
                foreach ($sectionXml in $sectiongroup.Section) {
                    [Section]$section = [Section]::new($sectionXml)
                    $section.ToString()
                }
            }
        }

        # Checks for any sections not placed in a sectiongroup
        foreach ($sectionXml in $notebook.Section) {
            [Section]$section = [Section]::new($sectionXml)
            $section.ToString()
        }
    }
}


$str = Main
Set-Content -Path "OneNote x Powershell\FULLREPORT.txt" -Value $str
$str