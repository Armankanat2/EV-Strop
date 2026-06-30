param(
    [Parameter(Mandatory = $true)]
    [string]$PresentationPath,

    [Parameter(Mandatory = $true)]
    [string]$Code
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

        if ($slideCode -ne $Code) { continue }

        Write-Output ("SlideIndex {0}" -f $slide.SlideIndex)
        Write-Output ("Code {0}" -f $slideCode)
        Write-Output ("ShapeCount {0}" -f $slide.Shapes.Count)

        foreach ($shape in $slide.Shapes) {
            $text = ""
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text.Replace("`r", " ").Replace("`n", " ")
            }

            Write-Output ("Id={0}; Type={1}; Left={2:N1}; Top={3:N1}; W={4:N1}; H={5:N1}; Text={6}" -f $shape.Id, $shape.Type, $shape.Left, $shape.Top, $shape.Width, $shape.Height, $text)
        }
    }
}
finally {
    $presentation.Close()
    $ppt.Quit()
}
