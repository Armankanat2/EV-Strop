$ErrorActionPreference = "Stop"

$presentationPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S029_theme22_preview_2026-06-30_v13_ev-group-techstyle-premium-subslides.pptx"

$msoTrue = -1
$msoFalse = 0
$ppMouseClick = 1
$ppActionHyperlink = 7
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$ppAlignLeft = 1
$ppAlignCenter = 2

# PowerPoint COM expects BGR integers.
$ColorWhite = 0xFFFFFF
$ColorGraphite = 0x111111
$ColorLightGray = 0xEBE7E5
$ColorDeepBlue = 0xB83800

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

function Find-CodeSlidesAll {
    param($presentation, [string]$code)

    $results = @()
    foreach ($slide in $presentation.Slides) {
        foreach ($shape in $slide.Shapes) {
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $text = $shape.TextFrame.TextRange.Text.Trim()
                if ($text -eq $code) {
                    $results += $slide
                    break
                }
            }
        }
    }

    return $results
}

function Try-GetShapeById {
    param($slide, [int]$id)

    foreach ($shape in $slide.Shapes) {
        if ($shape.Id -eq $id) {
            return $shape
        }
    }

    return $null
}

function Set-Rgb {
    param($holder, [int]$rgb)
    $holder.RGB = $rgb
}

function Set-FillSolid {
    param($shape, [int]$rgb)
    $shape.Fill.Visible = $msoTrue
    $shape.Fill.Solid()
    Set-Rgb $shape.Fill.ForeColor $rgb
}

function Set-Line {
    param($shape, [int]$rgb, [double]$weight = 1.0)
    $shape.Line.Visible = $msoTrue
    Set-Rgb $shape.Line.ForeColor $rgb
    $shape.Line.Weight = $weight
}

function Hide-Line {
    param($shape)
    $shape.Line.Visible = 0
}

function Apply-TextStyle {
    param(
        $shape,
        [double]$size,
        [int]$color = $ColorGraphite,
        [int]$align = $ppAlignLeft,
        [bool]$bold = $false
    )

    $shape.TextFrame.WordWrap = $msoTrue
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 6
    $shape.TextFrame.MarginBottom = 6

    $range = $shape.TextFrame.TextRange
    $range.Font.Name = "Arial"
    $range.Font.Size = $size
    $range.Font.Bold = $(if ($bold) { $msoTrue } else { 0 })
    Set-Rgb $range.Font.Color $color
    $range.ParagraphFormat.Alignment = $align
}

function Set-CodeStyle {
    param($shape)

    $fontColor = $ColorGraphite
    if ($shape.Fill.Visible -eq $msoTrue -and $shape.Fill.ForeColor.RGB -eq $ColorDeepBlue) {
        $fontColor = $ColorWhite
    }

    Apply-TextStyle $shape 11.0 $fontColor $ppAlignLeft $true
}

function Update-SlideCodeTexts {
    param($slide, [string]$newCode)

    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
            $text = $shape.TextFrame.TextRange.Text.Trim()
            if ($text -match '^S\d{3}(-P\d{2}|-PP\d{2}|-P56-01)?$') {
                $shape.TextFrame.TextRange.Text = $newCode
                Set-CodeStyle $shape
            }
        }
    }
}

function Set-SlideJump {
    param($shape, $targetSlide)

    $shape.ActionSettings($ppMouseClick).Action = $ppActionHyperlink
    $shape.ActionSettings($ppMouseClick).Hyperlink.Address = ""
    $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = "$($targetSlide.SlideID),$($targetSlide.SlideIndex), "
}

function Clear-SlideJump {
    param($shape)

    try {
        $shape.ActionSettings($ppMouseClick).Action = 0
        $shape.ActionSettings($ppMouseClick).Hyperlink.Address = ""
        $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = ""
    }
    catch {
    }
}

function Remove-GeneratedShapes {
    param($slide, [string]$prefix)

    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $slide.Shapes.Item($i)
        if ($shape.Name -like "$prefix*") {
            $shape.Delete()
        }
    }
}

function Add-TextBox {
    param(
        $slide,
        [string]$name,
        [double]$left,
        [double]$top,
        [double]$width,
        [double]$height,
        [string]$text,
        [double]$size,
        [int]$color = $ColorGraphite,
        [int]$align = $ppAlignLeft,
        [bool]$bold = $false
    )

    $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $shape.Name = $name
    $shape.Fill.Visible = 0
    $shape.Line.Visible = 0
    $shape.TextFrame.TextRange.Text = $text
    Apply-TextStyle $shape $size $color $align $bold
    return $shape
}

function Add-FittedPicture {
    param(
        $slide,
        [string]$name,
        [string]$imagePath,
        [double]$left,
        [double]$top,
        [double]$width,
        [double]$height
    )

    if (-not (Test-Path $imagePath)) {
        throw "Image not found: $imagePath"
    }

    Add-Type -AssemblyName System.Drawing | Out-Null
    $img = [System.Drawing.Image]::FromFile($imagePath)
    try {
        $sourceWidth = [double]$img.Width
        $sourceHeight = [double]$img.Height
    }
    finally {
        $img.Dispose()
    }

    $scale = [Math]::Min($width / $sourceWidth, $height / $sourceHeight)
    $finalWidth = $sourceWidth * $scale
    $finalHeight = $sourceHeight * $scale
    $finalLeft = $left + (($width - $finalWidth) / 2)
    $finalTop = $top + (($height - $finalHeight) / 2)

    $picture = $slide.Shapes.AddPicture($imagePath, $msoFalse, $msoTrue, $finalLeft, $finalTop, $finalWidth, $finalHeight)
    $picture.Name = $name

    return $picture
}

function Remove-RightVisualPictures {
    param($slide)

    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $slide.Shapes.Item($i)
        if (($shape.Type -eq 11 -or $shape.Type -eq 13) -and $shape.Left -ge 370 -and $shape.Top -ge 170) {
            $shape.Delete()
        }
    }
}

function Add-NavCard {
    param(
        $slide,
        [string]$name,
        [double]$left,
        [double]$top,
        [double]$width,
        [double]$height,
        [string]$title,
        [double]$fontSize = 10.6
    )

    $card = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $card.Name = "${name}_panel"
    Set-FillSolid $card $ColorWhite
    Set-Line $card $ColorLightGray 1.0

    $strip = $slide.Shapes.AddShape($msoShapeRectangle, ($left + 8), ($top + 6), ($width - 16), 3.0)
    $strip.Name = "${name}_strip"
    Set-FillSolid $strip $ColorDeepBlue
    Hide-Line $strip

    Add-TextBox $slide "${name}_text" ($left + 12) ($top + 9) ($width - 24) ($height - 16) $title $fontSize $ColorGraphite $ppAlignLeft $true | Out-Null
    return $card
}

function Update-PPContent {
    param(
        $slide,
        [string]$code,
        [string]$title,
        [string]$leftTitle,
        [string]$leftBody,
        [string]$rightTitle,
        [string]$rightCaption,
        [string]$imagePath
    )

    Update-SlideCodeTexts $slide $code

    $titleShape = Try-GetShapeById $slide 5
    if ($titleShape) {
        $titleShape.TextFrame.TextRange.Text = $title
        Apply-TextStyle $titleShape 20.5 $ColorGraphite $ppAlignLeft $true
    }

    $leftTitleShape = Try-GetShapeById $slide 7
    if ($leftTitleShape) {
        $leftTitleShape.TextFrame.TextRange.Text = $leftTitle
        Apply-TextStyle $leftTitleShape 12.0 $ColorGraphite $ppAlignLeft $true
    }

    $leftBodyShape = Try-GetShapeById $slide 8
    if ($leftBodyShape) {
        $leftBodyShape.TextFrame.TextRange.Text = $leftBody
        Apply-TextStyle $leftBodyShape 10.4 $ColorGraphite $ppAlignLeft $false
    }

    $rightTitleShape = Try-GetShapeById $slide 23
    if ($rightTitleShape) {
        $rightTitleShape.TextFrame.TextRange.Text = $rightTitle
        Apply-TextStyle $rightTitleShape 12.0 $ColorGraphite $ppAlignLeft $true
    }

    $rightCaptionShape = Try-GetShapeById $slide 10
    if ($rightCaptionShape) {
        $rightCaptionShape.TextFrame.TextRange.Text = $rightCaption
        Apply-TextStyle $rightCaptionShape 9.2 $ColorGraphite $ppAlignLeft $false
    }

    Remove-GeneratedShapes $slide "Codex_S039_Image"
    Remove-RightVisualPictures $slide

    $leftPanelShape = Try-GetShapeById $slide 6
    $leftTitleBandShape = Try-GetShapeById $slide 7
    $leftBodyShape = Try-GetShapeById $slide 8
    $leftTopLineShape = Try-GetShapeById $slide 14
    $leftStepShape = Try-GetShapeById $slide 15
    $leftKindShape = Try-GetShapeById $slide 16
    $leftDividerShape = Try-GetShapeById $slide 17

    $rightPanelShape = Try-GetShapeById $slide 9
    $rightTitleShape = Try-GetShapeById $slide 23
    $rightTopLineShape = Try-GetShapeById $slide 19
    $rightCodeShape = Try-GetShapeById $slide 20
    $rightStepShape = Try-GetShapeById $slide 21
    $rightKindShape = Try-GetShapeById $slide 22
    $rightDividerShape = Try-GetShapeById $slide 18
    $rightCaptionDividerShape = Try-GetShapeById $slide 25

    $leftPanelLeft = 53.3
    $leftPanelTop = 144.0
    $leftPanelWidth = 286.0
    $leftPanelHeight = 308.2
    $rightPanelLeft = 351.3
    $rightPanelTop = 144.0
    $rightPanelWidth = 527.9
    $rightPanelHeight = 318.0

    if ($leftPanelShape) {
        $leftPanelShape.Left = $leftPanelLeft
        $leftPanelShape.Top = $leftPanelTop
        $leftPanelShape.Width = $leftPanelWidth
        $leftPanelShape.Height = $leftPanelHeight
    }

    if ($leftTitleBandShape) {
        $leftTitleBandShape.Left = $leftPanelLeft
        $leftTitleBandShape.Top = $leftPanelTop + 24.5
        $leftTitleBandShape.Width = $leftPanelWidth
        $leftTitleBandShape.Height = 24.5
    }

    if ($leftTopLineShape) {
        $leftTopLineShape.Left = $leftPanelLeft
        $leftTopLineShape.Top = $leftPanelTop
        $leftTopLineShape.Width = $leftPanelWidth
    }

    if ($leftStepShape) {
        $leftStepShape.Left = $leftPanelLeft + 8.6
        $leftStepShape.Top = $leftPanelTop + 7.2
    }

    if ($leftKindShape) {
        $leftKindShape.Left = ($leftPanelLeft + $leftPanelWidth) - 72.0 - 8.6
        $leftKindShape.Top = $leftPanelTop + 7.2
    }

    if ($leftDividerShape) {
        $leftDividerShape.Left = $leftPanelLeft + 8.6
        $leftDividerShape.Top = $leftPanelTop + 41.8
        $leftDividerShape.Width = $leftPanelWidth - 17.2
    }

    if ($leftBodyShape) {
        $leftBodyShape.Left = $leftPanelLeft + 17.3
        $leftBodyShape.Top = $leftPanelTop + 66.2
        $leftBodyShape.Width = $leftPanelWidth - 30.2
        $leftBodyShape.Height = $leftPanelHeight - 82.0
    }

    if ($rightPanelShape) {
        $rightPanelShape.Left = $rightPanelLeft
        $rightPanelShape.Top = $rightPanelTop
        $rightPanelShape.Width = $rightPanelWidth
        $rightPanelShape.Height = $rightPanelHeight
    }

    if ($rightTopLineShape) {
        $rightTopLineShape.Left = $rightPanelLeft
        $rightTopLineShape.Top = $rightPanelTop
        $rightTopLineShape.Width = $rightPanelWidth
    }

    if ($rightCodeShape) {
        $rightCodeShape.Left = $rightPanelLeft + 8.6
        $rightCodeShape.Top = $rightPanelTop + 7.2
    }

    if ($rightStepShape) {
        $rightStepShape.Left = $rightPanelLeft + 97.2
        $rightStepShape.Top = $rightPanelTop + 7.2
    }

    if ($rightKindShape) {
        $rightKindShape.Left = ($rightPanelLeft + $rightPanelWidth) - 72.0 - 8.7
        $rightKindShape.Top = $rightPanelTop + 7.2
    }

    if ($rightDividerShape) {
        $rightDividerShape.Left = $rightPanelLeft + 11.5
        $rightDividerShape.Top = $rightPanelTop + 53.3
        $rightDividerShape.Width = $rightPanelWidth - 23.1
    }

    if ($rightTitleShape) {
        $rightTitleShape.Left = $rightPanelLeft + 14.0
        $rightTitleShape.Top = $rightPanelTop + 20.0
        $rightTitleShape.Width = $rightPanelWidth - 28.0
        $rightTitleShape.Height = 18.0
    }

    if ($rightCaptionShape) {
        $rightCaptionShape.Left = $rightPanelLeft + 14.0
        $rightCaptionShape.Width = $rightPanelWidth - 28.0
        $rightCaptionShape.Height = 12.0
        $rightCaptionShape.Top = ($rightPanelTop + $rightPanelHeight) - 16.0
    }

    if ($rightCaptionDividerShape) {
        $rightCaptionDividerShape.Left = $rightPanelLeft + 12.9
        $rightCaptionDividerShape.Top = $rightCaptionShape.Top - 6.0
        $rightCaptionDividerShape.Width = $rightPanelWidth - 25.8
    }

    if ($rightPanelShape -and $rightTitleShape -and $rightCaptionShape) {
        $imageLeft = $rightPanelShape.Left + 12
        $imageTop = $rightTitleShape.Top + $rightTitleShape.Height + 6
        $imageWidth = $rightPanelShape.Width - 24
        $imageHeight = ($rightCaptionShape.Top - 6) - $imageTop

        Add-FittedPicture $slide "Codex_S039_Image_$code" $imagePath $imageLeft $imageTop $imageWidth $imageHeight | Out-Null
    }

    foreach ($id in @(15,21)) {
        $shape = Try-GetShapeById $slide $id
        if ($shape -and $shape.HasTextFrame -eq -1) {
            $shape.TextFrame.TextRange.Text = "Шаг 05"
            Apply-TextStyle $shape 10.0 $ColorGraphite $ppAlignCenter $true
        }
    }

    foreach ($id in @(16,22)) {
        $shape = Try-GetShapeById $slide $id
        if ($shape -and $shape.HasTextFrame -eq -1) {
            $shape.TextFrame.TextRange.Text = "Опора"
            Apply-TextStyle $shape 10.0 $ColorGraphite $ppAlignCenter $true
        }
    }
}

$createdApp = $false
try {
    Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null
    $ppt = [Microsoft.VisualBasic.Interaction]::GetObject("", "PowerPoint.Application")
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

$slideP01 = Find-CodeSlide $presentation "S039-P01"
$pp01Slides = @(Find-CodeSlidesAll $presentation "S039-PP01")
$slidePP01 = $pp01Slides[0]
$extraPp01Slides = @()
if ($pp01Slides.Count -gt 1) {
    $extraPp01Slides = $pp01Slides | Select-Object -Skip 1
}
$slideS040 = Find-CodeSlide $presentation "S040"
$slideS039 = Find-CodeSlide $presentation "S039"
$slideS038P01 = Find-CodeSlide $presentation "S038-P01"
$slideS041 = Find-CodeSlide $presentation "S041"
$imageDir = "C:\Users\Дмитрий\Pictures\дпо стропальщик"

foreach ($code in @("S039-PP05","S039-PP04","S039-PP03","S039-PP02")) {
    $existing = Find-CodeSlideOptional $presentation $code
    if ($existing) {
        $existing.Delete()
    }
}

foreach ($extraSlide in ($extraPp01Slides | Sort-Object SlideIndex -Descending)) {
    $extraSlide.Delete()
}

Remove-GeneratedShapes $slideP01 "Codex_S039"

$leftBody = Try-GetShapeById $slideP01 8
if ($leftBody) {
    $leftBody.Delete()
}

$leftTitle = Try-GetShapeById $slideP01 7
if ($leftTitle) {
    $leftTitle.TextFrame.TextRange.Text = "Выберите тему складирования"
    Apply-TextStyle $leftTitle 12.0 $ColorGraphite $ppAlignLeft $true
}

foreach ($id in @(6,7,14,15,16,17)) {
    $shape = Try-GetShapeById $slideP01 $id
    if ($shape) {
        Clear-SlideJump $shape
    }
}

$cardSpecs = @(
    @{ Name = "Codex_S039_Card1"; Text = "требования к складированию грузов"; Code = "S039-PP01"; Top = 206.0; Height = 31.0 }
    @{ Name = "Codex_S039_Card2"; Text = "складирование труб"; Code = "S039-PP02"; Top = 244.0; Height = 31.0 }
    @{ Name = "Codex_S039_Card3"; Text = "складирование лесоматериалов"; Code = "S039-PP03"; Top = 282.0; Height = 31.0 }
    @{ Name = "Codex_S039_Card4"; Text = "складирование металопроката"; Code = "S039-PP04"; Top = 320.0; Height = 31.0 }
    @{ Name = "Codex_S039_Card5"; Text = "складирование ферм"; Code = "S039-PP05"; Top = 358.0; Height = 31.0 }
)

foreach ($spec in $cardSpecs) {
    Add-NavCard $slideP01 $spec.Name 70.6 $spec.Top 290.2 $spec.Height $spec.Text 10.3 | Out-Null
}

$insertAfter = $slidePP01
$createdSlides = @{}
foreach ($code in @("S039-PP02","S039-PP03","S039-PP04","S039-PP05")) {
    $newSlide = $slidePP01.Duplicate().Item(1)
    $newSlide.MoveTo($insertAfter.SlideIndex)
    $insertAfter = $newSlide
    $createdSlides[$code] = $newSlide
}

Update-PPContent $slidePP01 `
    "S039-PP01" `
    "Подподвал. Требования к складированию грузов" `
    "Что проверяем" `
    "• складируют только на подготовленное и ровное основание;`r`n• применяют подкладки и прокладки под форму груза;`r`n• не загромождают проходы, проезды и рабочую зону;`r`n• груз укладывают устойчиво, без риска переката, сползания и опрокидывания." `
    "Требования к складированию" `
    "Схема по общим требованиям к складированию грузов." `
    (Join-Path $imageDir "требования_к_складированию_грузов.png")

Update-PPContent $createdSlides["S039-PP02"] `
    "S039-PP02" `
    "Подподвал. Складирование труб" `
    "Что показать по трубам" `
    "• трубы укладывают на подкладки и упоры;`r`n• исключают самопроизвольное раскатывание;`r`n• высоту штабеля выбирают устойчивой;`r`n• между штабелями оставляют безопасные проходы." `
    "Складирование труб" `
    "Схема по безопасному складированию труб." `
    (Join-Path $imageDir "складирование_труб.png")

Update-PPContent $createdSlides["S039-PP03"] `
    "S039-PP03" `
    "Подподвал. Складирование лесоматериалов" `
    "Что показать по лесоматериалам" `
    "• лесоматериалы укладывают в устойчивые пакеты или штабели;`r`n• применяют прокладки и ограничители смещения;`r`n• исключают перекос и разваливание штабеля;`r`n• проходы и подходы к штабелю держат свободными." `
    "Складирование лесоматериалов" `
    "Схема по безопасному складированию лесоматериалов." `
    (Join-Path $imageDir "складирование_лесоматериалов.png")

Update-PPContent $createdSlides["S039-PP04"] `
    "S039-PP04" `
    "Подподвал. Складирование металопроката" `
    "Что показать по металопрокату" `
    "• металлопрокат укладывают по видам и профилю;`r`n• под пакет ставят прочные подкладки;`r`n• исключают смещение и скатывание элементов;`r`n• длиномер и листы размещают так, чтобы не было перекоса." `
    "Складирование металопроката" `
    "Схема по безопасному складированию металопроката." `
    (Join-Path $imageDir "складирование_металлопроката.png")

Update-PPContent $createdSlides["S039-PP05"] `
    "S039-PP05" `
    "Подподвал. Складирование ферм" `
    "Что показать по фермам" `
    "• фермы укладывают на рассчитанные опоры и подкладки;`r`n• положение должно исключать прогиб, перекос и опрокидывание;`r`n• между элементами выдерживают безопасный зазор;`r`n• расстроповку выполняют только после полной устойчивой установки." `
    "Складирование ферм" `
    "Схема по безопасному складированию ферм." `
    (Join-Path $imageDir "складирование_ферм.png")

$slidePP01 = Find-CodeSlide $presentation "S039-PP01"
$slidePP02 = Find-CodeSlide $presentation "S039-PP02"
$slidePP03 = Find-CodeSlide $presentation "S039-PP03"
$slidePP04 = Find-CodeSlide $presentation "S039-PP04"
$slidePP05 = Find-CodeSlide $presentation "S039-PP05"
$slideS039 = Find-CodeSlide $presentation "S039"
$slideS040 = Find-CodeSlide $presentation "S040"
$slideS041 = Find-CodeSlide $presentation "S041"

$baseIndex = $slideS038P01.SlideIndex + 1
$slideS039.MoveTo($baseIndex)
$slideP01.MoveTo($baseIndex + 1)
$slidePP01.MoveTo($baseIndex + 2)
$slidePP02.MoveTo($baseIndex + 3)
$slidePP03.MoveTo($baseIndex + 4)
$slidePP04.MoveTo($baseIndex + 5)
$slidePP05.MoveTo($baseIndex + 6)
$slideS040.MoveTo($baseIndex + 7)
$slideS041.MoveTo($baseIndex + 8)

$slidePP01 = Find-CodeSlide $presentation "S039-PP01"
$slidePP02 = Find-CodeSlide $presentation "S039-PP02"
$slidePP03 = Find-CodeSlide $presentation "S039-PP03"
$slidePP04 = Find-CodeSlide $presentation "S039-PP04"
$slidePP05 = Find-CodeSlide $presentation "S039-PP05"

$targets = @($slidePP01, $slidePP02, $slidePP03, $slidePP04, $slidePP05)
for ($i = 0; $i -lt $targets.Count; $i++) {
    $base = "Codex_S039_Card$($i+1)"
    foreach ($suffix in @("_panel","_strip","_text")) {
        $shape = $slideP01.Shapes.Item("$base$suffix")
        Set-SlideJump $shape $targets[$i]
    }
}

Set-SlideJump (Try-GetShapeById $slidePP01 11) $slideP01
Set-SlideJump (Try-GetShapeById $slidePP01 12) $slidePP02

Set-SlideJump (Try-GetShapeById $slidePP02 11) $slideP01
Set-SlideJump (Try-GetShapeById $slidePP02 12) $slidePP03

Set-SlideJump (Try-GetShapeById $slidePP03 11) $slideP01
Set-SlideJump (Try-GetShapeById $slidePP03 12) $slidePP04

Set-SlideJump (Try-GetShapeById $slidePP04 11) $slideP01
Set-SlideJump (Try-GetShapeById $slidePP04 12) $slidePP05

Set-SlideJump (Try-GetShapeById $slidePP05 11) $slideP01
Set-SlideJump (Try-GetShapeById $slidePP05 12) $slideS040

$presentation.Save()

if (-not $wasAlreadyOpen) {
    $presentation.Close()
}

if ($createdApp) {
    $ppt.Quit()
}
