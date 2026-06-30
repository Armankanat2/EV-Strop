$ErrorActionPreference = "Stop"

$SlideIndex = if ($args.Count -gt 0) { [int]$args[0] } else { 1 }

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1

$path = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-29_v55.pptx"
$pres = $ppt.Presentations.Open($path, -1, $false, $false)

$slide = $pres.Slides.Item($SlideIndex)
Write-Output ("SlideIndex {0}" -f $SlideIndex)
Write-Output ("SlideSize {0}x{1}" -f $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight)

foreach ($shape in $slide.Shapes) {
    $text = ""
    if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
        $text = $shape.TextFrame.TextRange.Text.Replace("`r", " ").Replace("`n", " ")
    }

    Write-Output ("Id={0}; Type={1}; Left={2:N1}; Top={3:N1}; W={4:N1}; H={5:N1}; Text={6}" -f $shape.Id, $shape.Type, $shape.Left, $shape.Top, $shape.Width, $shape.Height, $text)
}

$pres.Close()
$ppt.Quit()
