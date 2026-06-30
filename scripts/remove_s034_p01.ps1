$ErrorActionPreference = "Stop"

$presentationPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S029_theme22_preview_2026-06-30_v13_ev-group-techstyle-premium-subslides.pptx"

$msoTrue = -1
$ppMouseClick = 1
$ppActionNone = 0

function Find-CodeSlide {
    param($presentation, [string]$code)

    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text.Trim()
                if ($text -eq $code) {
                    return $slide
                }
            }
        }
    }

    throw "Slide with code '$code' not found."
}

function Find-CodeSlideOptional {
    param($presentation, [string]$code)

    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text.Trim()
                if ($text -eq $code) {
                    return $slide
                }
            }
        }
    }

    return $null
}

function Clear-SlideJump {
    param($shape)

    try {
        $shape.ActionSettings($ppMouseClick).Action = $ppActionNone
        $shape.ActionSettings($ppMouseClick).Hyperlink.Address = ""
        $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = ""
    }
    catch {
    }
}

function Remove-LinksToSlide {
    param($slide, [string]$targetSlideId)

    foreach ($shape in $slide.Shapes) {
        try {
            $subAddress = [string]$shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress
            if ($subAddress -like "$targetSlideId,*") {
                Clear-SlideJump $shape
            }
        }
        catch {
        }
    }
}

$createdApp = $false
try {
    $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
}
catch {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue
    $createdApp = $true
}

$presentation = $null
$wasAlreadyOpen = $false
foreach ($openPresentation in $ppt.Presentations) {
    if ($openPresentation.FullName -eq $presentationPath) {
        $presentation = $openPresentation
        $wasAlreadyOpen = $true
        break
    }
}

if (-not $presentation) {
    $presentation = $ppt.Presentations.Open($presentationPath, $false, $false, $false)
}

$slideS034 = Find-CodeSlide $presentation "S034"
$slideS034P01 = Find-CodeSlideOptional $presentation "S034-P01"

if ($slideS034P01) {
    Remove-LinksToSlide $slideS034 $slideS034P01.SlideID
    $slideS034P01.Delete()
    $presentation.Save()
}

if (-not $wasAlreadyOpen) {
    $presentation.Close()
}

if ($createdApp) {
    $ppt.Quit()
}
