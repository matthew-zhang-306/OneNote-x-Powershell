using namespace System.Collections.Generic

class HtmlCreator {
    [List[string]]$Tags
    [Indenter]$Body

    HtmlCreator() {
        $this.Tags = [List[string]]::new()
        $this.Body = [Indenter]::new()
    }

    AddBreak() {
        $this.Body += "<br>"
    }

    AddTag([string]$tagName, [string]$className) {
        
        [string]$tag = "<" + $tagName + " class='" + $className + "'>"
    }
    CloseTag() {

    }

    AddText([string]$text) {
        $this.Body += $text
    }
}