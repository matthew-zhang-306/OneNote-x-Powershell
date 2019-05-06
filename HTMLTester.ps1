using namespace System.Collections.Generic

class SomeObject {
    [string]$Phrase
    [int[]]$Lengths
    [string]$LengthsDisplay
    [int]$Hash
    [datetime]$Date
    [bool]$IsAwesome
    [string]$Inner
    
    SomeObject([string]$str, [int]$num) {
        $this.Phrase = $str
        
        $this.Lengths = @()
        foreach ($word in @($str -split " ")) {
            $this.Lengths += $word.Length
        }

        $this.LengthsDisplay = $this.Lengths -join ", "

        $this.Hash = $str.GetHashCode()

        $this.Date = (Get-Date).AddDays($num)

        $this.IsAwesome = $this.LengthsDisplay.GetHashCode() -gt 0

        $this.Inner = ($this.Lengths | ConvertTo-Html -Fragment)
    }
}




$Header = @"
<title>HTML Test</title>
<style>
* {font-family: 'Courier New', monospace; box-sizing: border-box;}
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #FFD700;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

$PreContent = @"
<h1>Some Objects</h1>
"@

[SomeObject[]]$objs = @(
    [SomeObject]::new("i wumbo you wumbo he she wumbo", 30),
    [SomeObject]::new("the snack that smiles back would actually probably just be oranges cut into slices", 40),
    [SomeObject]::new("rose is red violet is blue flag is win baba is you", 99),
    [SomeObject]::new("ZOOOOOOOOOOOOOOOOOOOOOM", 199)
)
$objs
" "

$html = $objs | ConvertTo-Html -Property Phrase, LengthsDisplay, Inner, Date, IsAwesome -As Table -Head $Header -PreContent $PreContent
Set-Content -Path "OneNote x Powershell\Reports\htmltest.html" -Value $html
$html
