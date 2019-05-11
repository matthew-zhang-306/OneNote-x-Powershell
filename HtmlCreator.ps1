using namespace System.Collections.Generic

class HtmlCreator {
    [List[string]]$Tags
    [string]$Body

    HtmlCreator() {
        $this.Tags = [List[string]]::new()
        $this.Body = ""
    }

    AddBreak() {
        $this.Body += "<br>"
    }
}