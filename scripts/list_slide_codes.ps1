param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath
)

$ErrorActionPreference = "Stop"

$fullPath = if ([System.IO.Path]::IsPathRooted($PresentationPath)) {
    $PresentationPath
} else {
    Join-Path $PWD $PresentationPath
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$presentation = $ppt.Presentations.Open($fullPath, $true, $false, $false)

try {
    foreach ($slide in $presentation.Slides) {
        $slideCode = ""
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text.Trim()
                if ($text -match '^S\d{3}(-P\d{2})?(-PP\d{2})?$') {
                    $slideCode = $text
                    break
                }
            }
        }

        Write-Output ("{0}: {1}" -f $slide.SlideIndex, $slideCode)
    }
}
finally {
    $presentation.Close()
    $ppt.Quit()
}
