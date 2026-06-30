$ErrorActionPreference = "Stop"

$presentationPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-29_v55.pptx"
$imagePath = Join-Path $PWD "assets/course-media/module-01-stropovka-gruzov/images/S019A_kreplenie_dlinnomer_25.png"

$msoTrue = -1
$msoFalse = 0
$ppMouseClick = 1
$ppActionHyperlink = 7
$ppAlignLeft = 1
$ppAlignCenter = 2

function Find-SlideByCode {
    param(
        $Presentation,
        [string]$Code
    )

    foreach ($slide in $Presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
                if ($shape.TextFrame.TextRange.Text.Trim() -eq $Code) {
                    return $slide
                }
            }
        }
    }

    throw "Slide with code '$Code' not found."
}

function Find-ShapeByName {
    param(
        $Slide,
        [string]$Name
    )

    foreach ($shape in $Slide.Shapes) {
        if ($shape.Name -eq $Name) {
            return $shape
        }
    }

    throw "Shape '$Name' not found on slide."
}

function Remove-GeneratedShapes {
    param(
        $Slide,
        [string]$Prefix
    )

    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Slide.Shapes.Item($i)
        if ($shape.Name -like "$Prefix*") {
            $shape.Delete()
        }
    }
}

function Set-SlideJump {
    param(
        $Shape,
        $TargetSlide
    )

    $action = $Shape.ActionSettings($ppMouseClick)
    $action.Action = $ppActionHyperlink
    $action.Hyperlink.Address = ""
    $action.Hyperlink.SubAddress = ($TargetSlide.SlideID.ToString() + "," + $TargetSlide.SlideIndex.ToString() + ",")
}

function Add-FittedPicture {
    param(
        $Slide,
        [string]$Name,
        [string]$Path,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height
    )

    Add-Type -AssemblyName System.Drawing | Out-Null
    $img = [System.Drawing.Image]::FromFile($Path)
    try {
        $sourceWidth = [double]$img.Width
        $sourceHeight = [double]$img.Height
    }
    finally {
        $img.Dispose()
    }

    $scale = [Math]::Min($Width / $sourceWidth, $Height / $sourceHeight)
    $finalWidth = $sourceWidth * $scale
    $finalHeight = $sourceHeight * $scale
    $finalLeft = $Left + (($Width - $finalWidth) / 2)
    $finalTop = $Top + (($Height - $finalHeight) / 2)

    $picture = $Slide.Shapes.AddPicture($Path, $msoFalse, $msoTrue, $finalLeft, $finalTop, $finalWidth, $finalHeight)
    $picture.Name = $Name
    return $picture
}

function Set-Text {
    param(
        $Shape,
        [string]$Text,
        [double]$FontSize = 16,
        [bool]$Bold = $false,
        [int]$Align = $ppAlignLeft
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { $msoTrue } else { $msoFalse })
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
}

if (-not (Test-Path $presentationPath)) {
    throw "Presentation not found: $presentationPath"
}

if (-not (Test-Path $imagePath)) {
    throw "Image not found: $imagePath"
}

Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null
$ppt = $null
$presentation = $null
$createdApp = $false

try {
    $ppt = [Microsoft.VisualBasic.Interaction]::GetObject("", "PowerPoint.Application")
}
catch {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue
    $createdApp = $true
}

foreach ($openPresentation in $ppt.Presentations) {
    if ($openPresentation.FullName -eq (Resolve-Path $presentationPath).Path) {
        $presentation = $openPresentation
        break
    }
}

if (-not $presentation) {
    $presentation = $ppt.Presentations.Open((Resolve-Path $presentationPath).Path, $msoFalse, $msoFalse, $msoFalse)
}

$slideS018 = Find-SlideByCode $presentation "S018"
$slideS019 = Find-SlideByCode $presentation "S019"
$slideS019P01 = Find-SlideByCode $presentation "S019-P01"
$slideS020 = Find-SlideByCode $presentation "S020"
$slideS021 = Find-SlideByCode $presentation "S021"

$slideS019A = $null
try {
    $slideS019A = Find-SlideByCode $presentation "S019A"
}
catch {
    $slideS019A = $null
}

if (-not $slideS019A) {
    $duplicateRange = $slideS019.Duplicate()
    $slideS019A = $duplicateRange.Item(1)
}

$slideS019A.MoveTo($slideS019.SlideIndex + 1)

$shapeTitle = Find-ShapeByName $slideS019A "TextBox 2"
$shapeSubtitle = Find-ShapeByName $slideS019A "TextBox 3"
$shapeCode = Find-ShapeByName $slideS019A "Rounded Rectangle 4"
$shapeLeftPanel = Find-ShapeByName $slideS019A "Rounded Rectangle 5"
$shapeLeftHeader = Find-ShapeByName $slideS019A "Rounded Rectangle 6"
$shapeLeftBody = Find-ShapeByName $slideS019A "TextBox 7"
$shapeLeftCaption = Find-ShapeByName $slideS019A "TextBox 9"
$shapeRightPanel = Find-ShapeByName $slideS019A "Rounded Rectangle 10"
$shapeRightHeader = Find-ShapeByName $slideS019A "Rounded Rectangle 11"
$shapeRightBody = Find-ShapeByName $slideS019A "TextBox 12"
$shapeBack = Find-ShapeByName $slideS019A "Rounded Rectangle 13"
$shapeNext = Find-ShapeByName $slideS019A "Rounded Rectangle 14"

Set-Text $shapeTitle "Правила строповки. Удавка на длинномере" 24 $true $ppAlignLeft
Set-Text $shapeSubtitle "Точки обвязки размещают симметрично: ориентир - около 25% длины от каждого края." 16 $false $ppAlignLeft
Set-Text $shapeCode "S019A" 16 $true $ppAlignCenter
$shapeCode.Width = 140

Set-Text $shapeLeftHeader "Схема 25% от края" 16 $true $ppAlignLeft
Set-Text $shapeLeftBody "" 16 $false $ppAlignLeft
$shapeLeftBody.Height = 1

Set-Text $shapeLeftCaption "Типовая схема строповки длинномерного груза удавкой: точки обвязки ставят на расстоянии 1/4L от краев." 14 $false $ppAlignLeft
$shapeLeftCaption.Left = 82.1
$shapeLeftCaption.Top = 394.0
$shapeLeftCaption.Width = 361.4
$shapeLeftCaption.Height = 50.0

Set-Text $shapeRightHeader "Что запомнить" 16 $true $ppAlignLeft
Set-Text $shapeRightBody @"
• Удавки на длинномерном грузе ставят симметрично.
• Рабочий ориентир - примерно 25% длины от каждого края.
• Такое расположение уменьшает прогиб, перекос и риск соскальзывания.
• Нельзя ставить удавки слишком близко к краям или стягивать их к центру.
"@ 16 $false $ppAlignLeft
$shapeRightBody.Left = 525.6
$shapeRightBody.Top = 182.0
$shapeRightBody.Width = 336.0
$shapeRightBody.Height = 220.0

Remove-GeneratedShapes $slideS019A "Codex_S019A_Image"
$picture = Add-FittedPicture $slideS019A "Codex_S019A_Image" $imagePath 78 176 368 208
$picture.Line.Visible = $msoFalse

Set-SlideJump (Find-ShapeByName $slideS019 "Rounded Rectangle 14") $slideS019A
Set-SlideJump (Find-ShapeByName $slideS020 "Rounded Rectangle 13") $slideS019A

Set-SlideJump $shapeBack $slideS019
Set-SlideJump $shapeNext $slideS020

Set-SlideJump (Find-ShapeByName $slideS019 "Rounded Rectangle 13") $slideS018
Set-SlideJump (Find-ShapeByName $slideS020 "Rounded Rectangle 14") $slideS021

$presentation.Save()

if ($createdApp) {
    $presentation.Close()
    $ppt.Quit()
}
