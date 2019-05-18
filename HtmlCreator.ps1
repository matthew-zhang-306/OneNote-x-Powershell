using namespace System.Collections.Generic

Import-Module Indenter -PassThru

class HtmlCreator {
    [List[string]]$Tags
    $Body

    HtmlCreator() {
        $this.Tags = [List[string]]::new()
        $this.Body = Get-NewIndenter
    }

    AddBreak() {
        $this.Body += "<br>"
    }

    AddTag([string]$tagName, [string]$className) {
        $this.Tags.Add($tagName)
        
        [string]$tag = "<" + $tagName + " class='" + $className + "'>"
        $this.Body += $tagName

        $this.Body.IncreaseIndent()
    }
    CloseTag() {
        if ($this.Tags.Count -gt 0) {
            $tag = $this.Tags[$this.Tags.Count - 1]
            $this.Tags.RemoveAt($this.Tags.Count - 1)

            $this.Body.DecreaseIndent()
            $this.Body += "</" + $tag + ">"
        }
    }

    AddText([string]$text) {
        $this.Body += $text
    }

    AddHtml([string]$html) {
        $this.Body += $html
    }

    [string]ToString() {
        return $this.Body.Print()
    }
}