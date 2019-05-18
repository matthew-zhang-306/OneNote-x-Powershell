using namespace System.Xml

Import-Module Rectangle -PassThru

class Ink {
    static [bool]$Debug = $false

    $Rect
    [string]$Text

    Ink([XmlElement]$ink, [bool]$isWord) {
        if ($isWord) {
            $this.Rect = Get-NewRectangle(-$ink.inkOriginX, -$ink.inkOriginY, $ink.width, $ink.height)
            $this.Text = "[Text]: " + $ink.recognizedText
        } else {
            $this.Rect = Get-NewRectangle($ink.Position.X, $ink.Position.Y, $ink.Size.Width, $ink.Size.Height)
            $this.Text = "[Drawing]"
        }
    }

    [string]ToString() {
        return $this.Text +
            $(if ($this.Text.Length -gt 0) { " " } else { "" }) +
            $(if ([Ink]::Debug) { $this.Rect.ToString() } else { "" })
    }
}

function Get-NewInk($ink, $isWord) {
  return [Ink]::new($ink, $isWord)
}

Export-ModuleMember -Function Get-NewInk