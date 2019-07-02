using namespace Microsoft.Office.Interop
using namespace System.Collections.Generic
using namespace System.Xml

# Everything is back in one file again because Powershell really dislikes classes being in separate files.


cls

<#
RECTANGLE CLASS

Stores information about the position of an axis-aligned rectangle.
#>
class Rectangle {
    [float]$X
    [float]$Y
    [float]$Width
    [float]$Height

    [float]$Left
    [float]$Right
    [float]$Top
    [float]$Bottom

    # Constructor using x, y, width, and height
    # Usage: $rect = [Rectangle]::new(0, 0, 10, 10)
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

    # Returns whether or not rectangle $r overlaps with this rectangle
    # Usage: $rect.Intersects($otherRect) <== $true, eg
    [bool]Intersects([Rectangle]$r) {
        return -not (($this.Left -gt $r.Right) -or ($r.Left -gt $this.Right) -or ($this.Top -gt $r.Bottom) -or ($r.Top -gt $this.Bottom))
    }

    # Returns the area of the rectangle
    # Usage: $rect.GetArea() <== 4.0, eg
    [float]GetArea() {
        return $this.Width * $this.Height
    }

    # Returns an output string describing the object
    # Usage: $rect.ToString() <== RECT[0, 0, 10, 10], eg
    [string]ToString() {
        return "RECT[" + $this.X + ", " + $this.Y + ", " + $this.Width + ", " + $this.Height + "]"
    }
}


<#
DATEHELPER CLASS

Contains static variables and objects to facilitate date calculations
#>
class DateHelper {
    # A set of valid weekdays and their associated order
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
    
    # The current time
    static [datetime]$Now = (Get-Date -Year 2019 -Month 6 -Day 11) # Comment out parameters to use the current date and not a debug time
    
    # The current date
    static [datetime]$Today = [DateHelper]::Now.Date

    # Returns whether two times are on the same date
    # Usage: [DateHelper]::IsSameDay($d1, $d2) <== $true, eg
    static [bool]IsSameDay([datetime]$date1, [datetime]$date2) {
        return $date1.Date.ToString() -eq $date2.Date.ToString()
    }

    # Returns whether a label is an existing weekday
    # Usage: [DateHelper]::IsValidWeekday("monday") <== $true
    static [bool]IsValidWeekday([string]$weekday) {
        return [DateHelper]::WeekdayMap.ContainsKey([DateHelper]::PascalCase($weekday))
    }

    # Returns the weekday on which a time falls
    # Usage: [DateHelper]::GetWeekday($d) <== "Monday", eg
    static [string]GetWeekday([datetime]$date) {
        return $date.DayOfWeek.ToString()
    }

    # Returns whether the given date is the same weekday as another date or a weekday string
    # Usage: [DateHelper]::IsSameWeekday($d1, $d2) <== $true, eg
    static [bool]IsSameWeekday([datetime]$date1, [datetime]$date2) {
        return [DateHelper]::GetWeekday($date1) -eq [DateHelper]::GetWeekday($date2)
    }
    static [bool]IsSameWeekday([datetime]$date, [string]$dateStr) {
        return [DateHelper]::GetWeekday($date) -eq [DateHelper]::PascalCase($dateStr)
    }

    # Converts a one-word string to pascal case
    # Used to convert raw weekday strings into formalized ones (eg "MONDAY" => "Monday") for comparison
    static [string]PascalCase([string]$str) {
        return $str.Substring(0, 1).ToUpper() + $str.Substring(1).ToLower()
    }
}


<#
INDENTER CLASS

Formats a report by inserting indentations before specific lines of text
#>
class Indenter {
    # The current stack of indentations
    [List[string]]$Indents

    # The lines in the output, with indents included
    [List[string]]$Lines

    # Constructor
    # Usage: $ind = [Indenter]::new()
    Indenter() {
        $this.ClearLines()
    }

    # Returns all of the lines in one string
    # Usage: Write-Host $ind.Print()
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
        # note to self: when splitting strings only '\r\n' works, but when joining strings only "`r`n" works. powershell even cares about which quote marks you use. the inconsistency is weird
        return $output -join "`r`n"
    }

    # Adds an indent to the output
    # Default indent is a 4-space tab
    # Usage: $ind.IncreaseIndent("  - ")
    IncreaseIndent() {
        $this.IncreaseIndent("    ")
    }
    IncreaseIndent([string]$indent) {
        $this.Indents.Add($indent)
    }

    # Removes an indent from the output
    # Usage: $ind.DecreaseIndent()
    DecreaseIndent() {
        if ($this.Indents.Count -gt 0) {
            $this.Indents.RemoveAt($this.Indents.Count - 1)
        }
    }

    # Returns the combined indent string which will preceed every line added to the output
    # Usage: $ind.GetCurrentIndent() + $str <== "            triple indented", eg
    [string]GetCurrentIndent() {
        return $this.Indents -join ""
    }

    # Adds one or more lines to the output, applying indentation
    # Usage: $ind.AddLine($str)
    AddLine([string]$line) {
        $this.AddLines($line -split '\r?\n')
    }
    AddLines([List[string]]$lines) {
        foreach($line in $lines) {
            $this.Lines.Add($this.GetCurrentIndent() + $line)
        }
    }

    # Overload for + operator which functions like AddLine()
    # Usage: $ind += ($str1 + " " + $str2)
    # Note: This function is mutating so in every situation use +=, not +!
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
    
    # Resets the object to a clean state
    # Usage: $ind.ClearLines()
    ClearLines() {
        $this.Indents = [List[string]]::new()
        $this.Lines = [List[string]]::new()
    }
}


<#
HTMLCREATOR CLASS

Formats HTML output, handling tags using an Indenter object
#>
class HtmlCreator {
    # The stack of currently open tags
    [List[string]]$Tags

    # The current HTML block
    [Indenter]$Body

    # Constructor
    # Usage: $html = [HtmlCreator]::new()
    HtmlCreator() {
        $this.Tags = [List[string]]::new()
        $this.Body = [Indenter]::new()
    }

    # Adds a single line break to the HTML
    # Usage: $html.AddBreak()
    AddBreak() {
        $this.Body += "<br>"
    }

    # Opens a tag in the HTML
    # Usage: $html.AddTag("div", "exampleDiv")
    AddTag([string]$tagName, [string]$className) {
        $this.Tags.Add($tagName)

        [string]$tag = "<" + $tagName + " class='" + $className + "'>"
        $this.Body += $tag

        $this.Body.IncreaseIndent()
    }

    # Creates a single HTML element with content between two tags
    # Works as a shortcut for AddTag => AddText => CloseTag
    # Usage: $html.AddElement("div, "exampleDiv", "<strong>Sample</strong> text")
    AddElement([string]$tagName, [string]$className, [string]$text) {
        $this.AddTag($tagName, $className)
        $this.AddText($text)
        $this.CloseTag()
    }

    # Closes the previously opened tag in the HTML
    # Usage: $html.CloseTag()
    CloseTag() {
        if ($this.Tags.Count -gt 0) {
            $tag = $this.Tags[$this.Tags.Count - 1]
            $this.Tags.RemoveAt($this.Tags.Count - 1)

            $this.Body.DecreaseIndent()
            $this.Body += "</" + $tag + ">"
        }
    }

    # Adds any text or HTML to the output
    # Does not validate the input so be sure that what goes in is proper HTML
    # Usage: $html.AddText("<strong>Sample</strong> text")
    AddText([string]$text) {
        $this.Body += $text
    }

    # Returns the output of the HTML
    # Usage: Set-Content -Path "out.html" -Value $html.ToString()
    [string]ToString() {
        return $this.Body.Print()
    }
}


<#
INK CLASS

Stores information about an ink mark on a notebook page
#>
class Ink {
    # Whether or not the ToString() output should include a lot of detail
    static [bool]$Debug = $false


    # The box which the ink mark occupies
    [Rectangle]$Rect

    # A label to identify the ink drawing
    [string]$Text

    # Constructor using the raw XML object and an indication of its type
    # $isWord should be $true if the element is one:InkWord and $false if the element is one:InkDrawing
    # Usage: $ink = [Ink]::new()
    Ink([XmlElement]$ink, [bool]$isWord) {
        if ($isWord) {
            $this.Rect = [Rectangle]::new(-$ink.inkOriginX, -$ink.inkOriginY, $ink.width, $ink.height)
            $this.Text = "[Text]: " + $ink.recognizedText
        } else {
            $this.Rect = [Rectangle]::new($ink.Position.X, $ink.Position.Y, $ink.Size.Width, $ink.Size.Height)
            $this.Text = "[Drawing]"
        }
    }

    # Returns a string detailing the ink
    [string]ToString() {
        return $this.Text +
            $(if ($this.Text.Length -gt 0) { " " } else { "" }) +
            $(if ([Ink]::Debug) { $this.Rect.ToString() } else { "" })
    }
}


<#
IMAGE CLASS

Stores information about an image (assumed to be an assigned book page) in a notebook page
#>
class Image {
    # The ink area to image area ratio which must be met for the work to be considered substantial
    static [float] $pageFillConstant = 0.005
    

    # The box which the image occupies
    [Rectangle]$Rect

    # The ink marks contained inside the image
    [List[Ink]]$Inks
    [float]$InkArea

    # Whether or not the page contains an adequate amount of marks
    [bool]$HasWork

    # Constructor using the raw XML object
    # Usage: $image = [Image]::new($imageXml)
    Image([XmlElement]$image) {
        $this.Rect = [Rectangle]::new($image.Position.X, $image.Position.Y, $image.Size.Width, $image.Size.Height)
    }

    # Finds out how much ink is on the page and determines work status
    # Usage: $image.SetInk($allInks)
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

    [string]FullReportHtml() {
        [HtmlCreator]$html = [HtmlCreator]::new()

        $imageDisplay = $this.Rect.ToString()
        if ($this.HasWork) {
            $imageDisplay += " (!)(has work)"
        }

        $html.AddElement("p", "fullReportImageSubheader", [string]$this.Inks.Count + " mark(s)")

        return $html.ToString()
    }
}


<#
PAGE CLASS

Stores information about a single notebook page
#>
class Page {
    # The number of days that must elapse before the page is considered inactive
    static [float]$ActiveThreshold = 3.0

    # Tag name to be used when there is no tag on the page
    static [string]$DefaultTagName = "# NoTag #"

    # Basic page information
    [string]$Name
    [string]$TagName
    [XmlElement]$Tag

    # The subject is obtained from the name of the parent Section
    [string]$Subject

    # Parent objects
    [Section]$Section
    [SectionGroup]$SectionGroup

    # Contained items
    [List[Image]]$Images
    [List[Ink]]$Inks

    # The time when the page was created
    [datetime]$CreationTime

    # The time when the page was last tagged by an instructor
    [datetime]$LastAssignedTime

    # The time when the page was previously edited
    [datetime]$LastModifiedTime

    # The date when the page was intended to be completed
    [datetime]$OriginalAssignmentDate

    # Formatted string for associating a date with the page
    [string]$DateDisplay

    # True when the page has been updated recently
    [bool]$Active
    
    # True when the page was updated
    [bool]$Changed

    # True when the page contains images with work
    [bool]$HasWork

    # True when the page contains no images
    [bool]$Empty

    # Constructor using the raw XML object, the content object from OneNote.Application, and the parent Section
    # Usage: $page = [Page]::new($pageXml, $content, $parentSection)
    Page([XmlElement]$page, [xml]$content, [Section]$section) {
        $this.Name = $page.Name
        $this.Section = $section
        $this.SectionGroup = $section.SectionGroup

        $this.Subject = $section.Subject

        $this.FetchTag($content)
        $this.FetchDates($page)

        $this.Inks = [List[Ink]]::new()
        $this.FetchInks($content)

        $this.Images = [List[Image]]::new()
        $this.FetchImages($content)

        $this.FetchStatus()
    }

    # Searches for tag information in the content
    # Used by the constructor
    FetchTag([xml]$content) {
        [XmlElement[]]$tags = $content.GetElementsByTagName("one:Tag")
        [XmlElement[]]$tagDefs = $content.GetElementsByTagName("one:TagDef")
        
        # Both items must eixst for there to be a tag
        if (($tags.Length -gt 0) -and ($tagDefs.Length -gt 0)) {
            $this.Tag = $tags[0]
            $this.TagName = $tagDefs[0].Name
        }
        else {
            $this.TagName = [Page]::DefaultTagName
        }
    }

    # Searches for date information in the xml
    # Used by the constructor
    FetchDates([XmlElement]$page) {
        $this.CreationTime = [datetime]$page.dateTime
        $this.LastModifiedTime = [datetime]$page.lastModifiedTime
        
        if ($this.TagName -eq [Page]::DefaultTagName) {
            $this.LastAssignedTime = [datetime]$this.CreationTime
        } else {
            $this.LastAssignedTime = [datetime]$this.Tag.creationDate
        }

        if ([DateHelper]::IsValidWeekday($this.SectionGroup.Name)) {
            $this.OriginalAssignmentDate = $this.CreationTime.Date
            while (-not [DateHelper]::IsSameWeekday($this.OriginalAssignmentDate, $this.SectionGroup.Name)) {
                $this.OriginalAssignmentDate = $this.OriginalAssignmentDate.AddDays(1)
            }
        }
        else {
            $this.OriginalAssignmentDate = $this.LastAssignedTime
        }

        # Set date string
        $this.DateDisplay = $this.OriginalAssignmentDate.ToString('MM/dd/yyyy')
    }

    # Searches for ink items in the content
    # Used by the constructor
    FetchInks([xml]$content) {
        # Check for ink drawings
        foreach ($ink in $content.GetElementsByTagName("one:InkDrawing")) {
            $this.Inks.Add([Ink]::new($ink, $false))
        }
        # Check for ink words
        foreach ($ink in $content.GetElementsByTagName("one:InkWord")) {
            $this.Inks.Add([Ink]::new($ink, $true))
        }
    }

    # Searches for image items in the content
    # Used by the constructor
    FetchImages([xml]$content) {
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
    }

    # Calculates status variables for the page using formulas
    FetchStatus() {
        $this.Active = $this.LastModifiedTime -gt [DateHelper]::Now.AddDays(-1 * [Page]::ActiveThreshold)
        $this.Changed = $this.LastModifiedTime -gt $this.LastAssignedTime
        $this.HasWork = $this.Images.Where({$_.HasWork -eq $true}).Count -gt 0
        $this.Empty = $this.Images.Count -eq 0
    }

    # Returns a string containing basic page information
    # Usage: Write-Host $page.ToString()
    [string]ToString() {
        return "PAGE: " + $this.Section.Notebook.Name.PadRight(40) + " | " + $this.Section.Name.PadRight(40) + " | " + $this.Name
    }


    [string]FullReport() {
        $indenter = [Indenter]::new()

        $statusDisplay = $this.DateDisplay
        if ($this.NeedsGrading) {
            $statusDisplay += " (!)(needs grading)"
        }
        elseif ($this.Changed) {
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

    [string]FullReportHtml() {
        [HtmlCreator]$html = [HtmlCreator]::new()

        $html.AddElement("p", "fullReportPageHeader", $this.Name)

        $html.AddTag("ul", "fullReportPageInfoList")

        $html.AddElement("li", "fullReportPageInfoDateItem", $this.DateDisplay)
        if ($this.TagName -ne [Page]::DefaultTagName) {
            $html.AddElement("li", "fullReportPageInfoTagItem", $this.TagName)
        }

        $html.AddTag("li", "fullReportPageInfoImageCountItem")
        $html.AddText("Pages:")
        $html.AddTag("ol", "fullReportImageList")
        foreach ($image in $this.Images) {
            $html.AddElement("li", "fullReportImageItem", $image.FullReportHtml())
        }
        $html.CloseTag()
        $html.CloseTag()

        $html.CloseTag()

        return $html.ToString()
    }
}


<#
SECTION CLASS

Stores information about a notebook section
#>
class Section {
    # Basic information
    [string]$Name
    [bool]$Deleted

    # Subject is obtained from the section name
    [string]$Subject

    # Contained pages
    [List[Page]]$Pages

    # Parent objects
    [SectionGroup]$SectionGroup
    [Notebook]$Notebook

    # Constructor using the raw XML object and the parent object
    # Use the $sectiongroup version unless there is no parent section group
    # Usage: $section = [Section]::new($sectionXml, $parentSectionGroup)
    Section([XmlElement]$section, [SectionGroup]$sectiongroup) {
        $this.Init($section, $sectiongroup.Notebook)
        $this.SectionGroup = $sectiongroup

        $this.Pages = [List[Page]]::new()
        $this.FetchPages($section)
    }
    Section([XmlElement]$section, [Notebook]$notebook) {
        $this.Init($section, $notebook)
        $this.CheckForSubject()
    }
    Init([XmlElement]$section, [Notebook]$notebook) {
        $this.Name = $section.Name
        $this.Deleted = $section.IsInRecycleBin
        $this.Notebook = $notebook
    }

    # Sees if the section name contains subject information and if so notifies the parent notebook
    # Used by the constructor
    CheckForSubject() {
        foreach ($subject in [Notebook]::AllSubjects) {
            if ($this.Name.ToLower().Contains($subject.ToLower())) {
                $this.Notebook.AddSubject($subject)
                $this.Subject = $subject
            }
        }
    }

    # Searches for contained pages in the xml
    # Used by the constructor
    FetchPages([XmlElement]$section) {
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
        if ($this.Deleted) {
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

    [string]FullReportHtml() {
        [HtmlCreator]$html = [HtmlCreator]::new()

        $html.AddElement("p", "fullReportSectionHeader", $this.Name)
        foreach ($page in $this.Pages) {
            $html.AddElement("div", "fullReportPageItem", $page.FullReportHtml())
        }

        return $html.ToString()
    }
}


<#
SECTIONGROUP CLASS

Stores information about a notebook section group
#>
class SectionGroup {
    # Basic information
    [string]$Name

    # Contained sections
    [List[Section]]$Sections

    # Parent object
    [Notebook]$Notebook

    # Constructor using the raw XML object and the parent notebook
    # Usage: $sectiongroup = [SectionGroup]::new($sectiongroupXml, $parentNotebook)
    SectionGroup([XmlElement]$sectiongroup, [Notebook]$notebook) {
        $this.Name = $sectiongroup.Name
        $this.Notebook = $notebook
        
        $this.Sections = [List[Section]]::new()
        $this.FetchSections($sectiongroup)
    }

    # Searches for contained sections
    # Used by the constructor
    FetchSections([XmlElement]$sectiongroup) {
        foreach ($sectionXml in $sectiongroup.Section) {
            $this.Sections.Add([Section]::new($sectionXml, $this))
        }
    }


    [string]FullReport() {
        $indenter = [Indenter]::new()
        
        # Header print
        $sectionDisplay = "# SectionGroup: " + $this.Name + " #"
        $indenter += $sectionDisplay

        # Section print
        $indenter.IncreaseIndent()
        foreach ($section in $this.Sections) {
            $indenter += $section.FullReport()
        }
        $indenter.DecreaseIndent()

        return $indenter.Print()
    }

    [string]FullReportHtml() {
        [HtmlCreator]$html = [HtmlCreator]::new()

        foreach ($section in $this.Sections) {
            $html.AddElement("div", "fullReportSectionItem", $section.FullReportHtml())
        }

        return $html.ToString()
    }
}



<#
NOTEBOOK CLASS

Stores information about a OneNote notebook
#>
class Notebook {
    # Array containing all assignable subject names
    static [string[]]$AllSubjects = "Math", "Reading", "Grammar"

    # Basic information
    [string]$Name
    [bool]$Deleted

    # Child items
    [List[Section]]$Sections
    [List[SectionGroup]]$SectionGroups

    # Subjects which are assigned to this particular notebook
    [List[string]]$Subjects

    # Constructor using the raw XML object
    # Usage: $notebook = [Notebook]::new($notebookXml)
    Notebook([XmlElement]$notebook) {
        $this.Name = $notebook.Name
        $this.Deleted = $notebook.IsInRecycleBin

        # Debug log xml
        # Set-Content -Path "log.txt" -Value $notebook.InnerXml

        $this.SectionGroups = [List[SectionGroup]]::new()
        $this.FetchSectionGroups($notebook)

        $this.Sections = [List[Section]]::new()
        $this.FetchSections($notebook)

        if ($this.Name -eq "Sai") {
            
        }
    }

    # Searches for contained section groups
    # Used by the constructor
    FetchSectionGroups([XmlElement]$notebook) {
        foreach ($sectiongroup in $notebook.SectionGroup) {
            if ($sectiongroup.isRecycleBin -ne $true) {
                $this.SectionGroups.Add([SectionGroup]::new($sectiongroup, $this))
            }
        }
    }

    # Updates the sections list to include the child items of the sectiongroup list, and goes through any sections placed outside of section groups
    FetchSections([XmlElement]$notebook) {
        foreach ($sectiongroup in $this.SectionGroups) {
            $this.Sections.AddRange($sectiongroup.Sections)  
        }

        foreach ($section in $notebook.Section) {
            [Section]::new($section, $this)
        }
    }

    # Adds a subject string to the subject list but only if it is valid and not there
    AddSubject([string]$subject) {
        if ($this.Subjects -eq $null) {
            $this.Subjects = [List[string]]::new()
        }
        if ($this.Subjects.IndexOf($subject) -lt 0) {
            $this.Subjects.Add($subject)
        }
    }

    # Returns a filtered list of contained pages that satisfy a given check function
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

    # Returns true if at least one of the contained pages satisfies the given check function
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

    # Status reports
    [List[Page]]GetUngradedPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.Changed -and $p.HasWork})
    }
    [List[Page]]GetInactivePages() {
        return $this.GetPagesWhere({param([Page]$p) -not $p.Active})
    }
    [List[Page]]GetEmptyPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.Empty})
    }
    [List[Page]]GetUnreviewedPages() {
        return $this.GetPagesWhere({param([Page]$p) $p.TagName -like "*REVIEW*"})
    }


    [string]MissingAssignmentReport([datetime]$date) {
        [Indenter]$indenter = [Indenter]::new()

        foreach ($subject in $this.Subjects) {
            $has = $this.HasPagesWhere({param([Page]$p) (-not $p.Empty) -and ($p.Subject.ToLower() -eq $subject.ToLower()) -and ([DateHelper]::IsSameDay($p.OriginalAssignmentDate, $date))})
            if (-not $has) {
                $indenter += ($this.Name + " - " + $subject)
            }
        }

        return $indenter.Print()
    }
    
    [string]MissingAssignmentReportHtml([datetime]$date) {
        [HtmlCreator]$html = [HtmlCreator]::new()

        foreach ($subject in $this.Subjects) {
            $has = $this.HasPagesWhere({param([Page]$p) (-not $p.Empty) -and ($p.Subject -eq $subject) -and ([DateHelper]::IsSameDay($p.OriginalAssignmentDate, $date))})
            if (-not $has) {
                $html.AddElement("li", "missingAssignmentStudentItem", $this.Name + " " + $subject)
            }
        }

        return $html.ToString()
    }

    [string]FullReport() {
        $indenter = [Indenter]::new()

        $indenter += " ", $this.Name, "-------------------"
        foreach ($sectiongroup in $this.SectionGroups) {
            $indenter += $sectiongroup.FullReport()
        }

        return $indenter.Print()
    }

    [string]FullReportHtml() {
        [HtmlCreator]$html = [HtmlCreator]::new()

        $html.AddTag("div", "fullReportNotebookContainer")
        
        $html.AddElement("p", "fullReportNotebookName", $this.Name)

        $html.AddTag("div", "fullReportSectionTableContainer")
        $html.AddTag("table", "fullReportSectionTable")
        
        $html.AddTag("tr", "fullReportSectionGroupHeaderRow")
        foreach ($sectiongroup in $this.SectionGroups) {
            $html.AddElement("th", "fullReportSectionGroupCellHeader", $sectiongroup.Name)
        }
        $html.CloseTag()

        $html.AddTag("tr", "fullReportSectionGrouopRow")
        foreach ($sectiongroup in $this.SectionGroups) {
            $html.AddElement("th", "fullReportSectionGroupCellItem", $sectiongroup.FullReportHtml())
        }
        $html.CloseTag()

        $html.CloseTag()
        $html.CloseTag()

        $html.CloseTag()

        return $html.ToString()
    }
}


<#
MAIN CLASS

Contains the main report generating functionality of the script
#>
class Main {
    static [string]$Path = "Reports\"
    static [string]$Style

    static [int]$MissingAssignmentLookahead = 7

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
        [HtmlCreator]$html = [HtmlCreator]::new()

        $html.AddTag("div", "fullReportContainer")
        foreach ($notebook in $this.Notebooks) {
            $html.AddText($notebook.FullReportHtml())
            $html.AddBreak()
        }
        $html.CloseTag()

        Set-Content -Path ("Reports\Html\FullReport.html") -Value $html.ToString()
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
        [HtmlCreator]$html = [HtmlCreator]::new()

        $html.AddTag("table", "statusReportTable")
        
        $html.AddTag("tr", "statusReportHeaderRow")
        $html.AddElement("th", "statusReportHeaderNotebook", "Notebook")
        $html.AddElement("th", "statusReportHeaderSectionGroup", "Section Group")
        $html.AddElement("th", "statusReportHeaderSection", "Section")
        $html.AddElement("th", "statusReportHeaderPage", "Page")
        $html.AddElement("th", "statusReportHeaderTag", "Tag")
        $html.CloseTag()

        foreach ($notebook in $this.Notebooks) {
            [List[Page]]$list = $func.Invoke($notebook)
            if ($list -eq $null) { continue }

            foreach ($page in $func.Invoke($notebook)) {
                $html.AddTag("tr", "statusReportPageRow")
                $html.AddElement("td", "statusReportPageNotebook", $page.Section.Notebook.Name)
                $html.AddElement("td", "statusReportPageSectionGroup", $page.SectionGroup.Name)
                $html.AddElement("td", "statusReportPageSection", $page.Section.Name)
                $html.AddElement("td", "statusReportPage", $page.Name)
                $html.AddElement("td", "statusReportPageTag", $page.TagName)
                $html.CloseTag()
            }
        }

        $html.CloseTag()

        Set-Content -Path ("Reports\Html\" + $name + ".html") -Value $html
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
        for([int]$i = 0; $i -lt [Main]::MissingAssignmentLookahead; $i++) {
            [datetime]$date = [DateHelper]::Today.AddDays($i + $sundayskip)
            if ([DateHelper]::IsSameWeekday($date, "SUNDAY")) {
                # No assignments on sundays: skip this day
                $sundayskip += 1
                $date = $date.AddDays(1)
            }

            $indenter += ($date.ToString().Substring(0, $date.ToString().IndexOf(" ")) + " missing:")
            $indenter.IncreaseIndent("    - ")

            foreach ($notebook in $this.Notebooks) {
                [string]$nOut = $notebook.MissingAssignmentReport($date)
                if ($nOut.Length -gt 0) {
                    $indenter += $nOut
                }
            }

            $indenter.DecreaseIndent()
            $indenter += " "
        }
        
        $str = $indenter.Print()
        Set-Content -Path ([Main]::Path + "MISSING ASSIGNMENT REPORT.txt") -Value $str
        return $str
    }

    MissingAssignmentReportHtml() {
        [HtmlCreator]$html = [HtmlCreator]::new()

        $html.AddTag("div", "missingAssignmentContainer")

        $sundayskip = 0
        for([int]$i = 0; $i -lt [Main]::MissingAssignmentLookahead; $i++) {
            $html.AddTag("div", "missingAssignmentDayContainer")

            [datetime]$date = [DateHelper]::Today.AddDays($i + $sundayskip)
            if ([DateHelper]::IsSameWeekday($date, "SUNDAY")) {
                # No assignments on sundays: skip this day
                $sundayskip += 1
                $date = $date.AddDays(1)
            }
            
            $html.AddElement("p", "missingAssignmentDayHeader", $date.ToString().Substring(0, $date.ToString().IndexOf(" ")))
            $html.AddElement("p", "missingAssignmentDaySubheader", "Students missing assignments:")

            $html.AddTag("ul", "missingAssignmentStudentList")
            foreach ($notebook in $this.Notebooks) {
                [string]$nOut = $notebook.MissingAssignmentReportHtml($date)
                if ($nOut.Length -gt 0) {
                    $html.AddText($nOut)
                }
            }
            $html.CloseTag()

            $html.CloseTag()
            $html.AddBreak()
        }

        $html.CloseTag()

        Set-Content -Path ("Reports\Html\MissingAssignmentReport.html") -Value $html.ToString()
    }
    
    
    UploadHtml() {
        try {
            $cred = Get-Content -Path "config.txt"
            $ip = $cred[0]
            $user = $cred[1]
            $pass = $cred[2]
            $onlinePath = $cred[3]

            $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
                Protocol = [WinSCP.Protocol]::Ftp
                HostName = $ip
                UserName = $user
                Password = $pass
            }

            $session = New-Object WinSCP.Session

            try {
                $session.Open($sessionOptions)

                $transferOptions = New-Object WinSCP.TransferOptions
                $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary

                # Add the new files
                $transferResult = $session.PutFiles((Get-Location).Path + "\Reports\Html", $onlinePath, $False, $transferOptions)
                $transferResult.Check()

                foreach ($transfer in $transferResult.Transfers)
                {
                    Write-Host "Upload of $($transfer.FileName) succeeded"
                }
            }
            finally {
                $session.Dispose()
            }
        }
        catch {
            Write-Host "Error uploading: $($_.Exception.Message)"
        }
    }
}


Function Main() {
    [Main]$main = [Main]::new()
    
    $main.FullReport()
    $main.FullReportHtml()
    " "
    " "
    $main.StatusReports()
    $main.StatusReportsHtml()
    " "
    " "
    $main.MissingAssignmentReport()
    $main.MissingAssignmentReportHtml()
    " "
    " "
    $main.UploadHtml()
}

Main