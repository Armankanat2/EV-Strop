$ErrorActionPreference = "Stop"

$presentationPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S029_theme22_preview_2026-06-30_v13_ev-group-techstyle-premium-subslides.pptx"

$msoTrue = -1
$ppMouseClick = 1
$ppActionHyperlink = 7
$msoShapeRectangle = 1
$ppAlignLeft = 1
$ppAlignCenter = 2

# PowerPoint COM expects BGR integers.
$ColorWhite = 0xFFFFFF
$ColorGraphite = 0x111111
$ColorMuted = 0x63554B
$ColorLightGray = 0xEBE7E5
$ColorSoftPanel = 0xF7F5F4
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

function Set-SlideJump {
    param($shape, $targetSlide)

    $shape.ActionSettings($ppMouseClick).Action = $ppActionHyperlink
    $shape.ActionSettings($ppMouseClick).Hyperlink.Address = ""
    $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = "$($targetSlide.SlideID),$($targetSlide.SlideIndex), "
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

function Update-SlideCodeTexts {
    param($slide, [string]$newCode)

    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
            $text = $shape.TextFrame.TextRange.Text.Trim()
            if ($text -match '^S\d{3}(-P\d{2})?(-PP\d{2})?$') {
                $shape.TextFrame.TextRange.Text = $newCode
                Apply-TextStyle $shape 11.0 $ColorGraphite $ppAlignLeft $true
            }
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

function Add-NavCard {
    param(
        $slide,
        [string]$name,
        [double]$left,
        [double]$top,
        [double]$width,
        [double]$height,
        [string]$title
    )

    $card = $slide.Shapes.AddShape($msoShapeRectangle, $left, $top, $width, $height)
    $card.Name = "${name}_panel"
    Set-FillSolid $card $ColorWhite
    Set-Line $card $ColorLightGray 1.0

    $strip = $slide.Shapes.AddShape($msoShapeRectangle, $left, $top, $width, 3.0)
    $strip.Name = "${name}_strip"
    Set-FillSolid $strip $ColorDeepBlue
    Hide-Line $strip

    $text = Add-TextBox $slide "${name}_text" ($left + 10) ($top + 10) ($width - 20) ($height - 16) $title 12.4 $ColorGraphite $ppAlignLeft $true
    return $card
}

function Build-TwoPanelFooter {
    param(
        $slide,
        [string]$title,
        [string]$code,
        [string]$leftTitle,
        [string]$leftBody,
        [string]$rightTitle,
        [string]$rightBody
    )

    Update-SlideCodeTexts $slide $code

    $titleShape = Try-GetShapeById $slide 5
    if ($titleShape) {
        $titleShape.TextFrame.TextRange.Text = $title
        Apply-TextStyle $titleShape 22.0 $ColorGraphite $ppAlignLeft $true
    }

    $codeMeta = Try-GetShapeById $slide 18
    if ($codeMeta) {
        Apply-TextStyle $codeMeta 10.0 $ColorGraphite $ppAlignLeft $true
    }

    Remove-GeneratedShapes $slide "Codex_S031P"

    $bodyShape = Try-GetShapeById $slide 10
    if ($bodyShape) {
        $bodyShape.Delete()
    }

    $panel = Try-GetShapeById $slide 6
    if ($panel) {
        $panel.Left = 53.3
        $panel.Top = 144.0
        $panel.Width = 825.8
        $panel.Height = 308.2
        Set-FillSolid $panel $ColorWhite
        Set-Line $panel $ColorLightGray 1.15
    }

    $topStrip = Try-GetShapeById $slide 14
    if ($topStrip) {
        $topStrip.Left = 53.3
        $topStrip.Top = 144.0
        $topStrip.Width = 825.8
        $topStrip.Height = 2.9
        Set-FillSolid $topStrip $ColorDeepBlue
        Hide-Line $topStrip
    }

    $innerLeft = 71.0
    $innerTop = 183.0
    $innerW = 377.0
    $innerH = 236.0
    $gap = 18.0
    $innerRight = $innerLeft + $innerW + $gap

    $leftPanel = $slide.Shapes.AddShape($msoShapeRectangle, $innerLeft, $innerTop, $innerW, $innerH)
    $leftPanel.Name = "Codex_S031P_LeftPanel"
    Set-FillSolid $leftPanel $ColorSoftPanel
    Set-Line $leftPanel $ColorLightGray 1.0

    $leftStrip = $slide.Shapes.AddShape($msoShapeRectangle, $innerLeft, $innerTop, $innerW, 3.0)
    $leftStrip.Name = "Codex_S031P_LeftStrip"
    Set-FillSolid $leftStrip $ColorDeepBlue
    Hide-Line $leftStrip

    $rightPanel = $slide.Shapes.AddShape($msoShapeRectangle, $innerRight, $innerTop, $innerW, $innerH)
    $rightPanel.Name = "Codex_S031P_RightPanel"
    Set-FillSolid $rightPanel $ColorWhite
    Set-Line $rightPanel $ColorLightGray 1.0

    $rightStrip = $slide.Shapes.AddShape($msoShapeRectangle, $innerRight, $innerTop, $innerW, 3.0)
    $rightStrip.Name = "Codex_S031P_RightStrip"
    Set-FillSolid $rightStrip $ColorDeepBlue
    Hide-Line $rightStrip

    Add-TextBox $slide "Codex_S031P_LeftTitle" ($innerLeft + 12) ($innerTop + 12) ($innerW - 24) 22 $leftTitle 12.0 $ColorGraphite $ppAlignLeft $true | Out-Null
    Add-TextBox $slide "Codex_S031P_LeftBody" ($innerLeft + 12) ($innerTop + 38) ($innerW - 24) ($innerH - 48) $leftBody 10.6 $ColorGraphite $ppAlignLeft $false | Out-Null

    Add-TextBox $slide "Codex_S031P_RightTitle" ($innerRight + 12) ($innerTop + 12) ($innerW - 24) 22 $rightTitle 12.0 $ColorGraphite $ppAlignLeft $true | Out-Null
    Add-TextBox $slide "Codex_S031P_RightBody" ($innerRight + 12) ($innerTop + 38) ($innerW - 24) ($innerH - 48) $rightBody 10.6 $ColorGraphite $ppAlignLeft $false | Out-Null
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

$slideS031 = Find-CodeSlide $presentation "S031"
$slideP01 = Find-CodeSlide $presentation "S031-P01"
$slideS032 = Find-CodeSlide $presentation "S032"

foreach ($code in @("S031-P05","S031-P04","S031-P03","S031-P02")) {
    $existing = Find-CodeSlideOptional $presentation $code
    if ($existing) {
        $existing.Delete()
    }
}

# Rebuild S031 left area as five separate clickable cards.
Remove-GeneratedShapes $slideS031 "Codex_S031"

foreach ($id in @(6,7,8,14,17)) {
    $shape = Try-GetShapeById $slideS031 $id
    if ($shape) {
        $shape.Delete()
    }
}

$cardSpecs = @(
    @{ Text = "Уточни задание"; Code = "S031-P01" }
    @{ Text = "Проверь массу и характер груза"; Code = "S031-P02" }
    @{ Text = "Проверь место работы"; Code = "S031-P03" }
    @{ Text = "Проверь схемы и таблицы"; Code = "S031-P04" }
    @{ Text = "Осмотри приспособления и тару"; Code = "S031-P05" }
)

$cardLeft = 53.3
$cardTop = 158.0
$cardWidth = 320.4
$cardHeight = 46.0
$cardGap = 11.0

for ($i = 0; $i -lt $cardSpecs.Count; $i++) {
    $spec = $cardSpecs[$i]
    $y = $cardTop + ($i * ($cardHeight + $cardGap))
    $card = Add-NavCard $slideS031 "Codex_S031_Card$($i+1)" $cardLeft $y $cardWidth $cardHeight $spec.Text
}

# Duplicate P01 four times and rebuild them.
$insertAfter = $slideP01
$newSlides = @{}
$footerSpecs = @(
    @{
        Code = "S031-P02"
        Title = "Подвал. Проверь массу и характер груза"
        LeftTitle = "Что проверить"
        LeftBody = "• Масса груза по документам, маркировке или паспорту.`r`n• Фактический вес не должен быть предположением.`r`n• Если масса не подтверждена, работу начинать нельзя."
        RightTitle = "Характер груза"
        RightBody = "• Габариты и форма груза.`r`n• Центр тяжести и устойчивость.`r`n• Выступающие части, тара, маркировка, особые условия перемещения."
    }
    @{
        Code = "S031-P03"
        Title = "Подвал. Проверь место работы"
        LeftTitle = "Рабочая площадка"
        LeftBody = "• Освещенность и видимость.`r`n• Опасная зона и сигнальное ограждение.`r`n• Отсутствие посторонних людей в зоне работ."
        RightTitle = "Маршрут и укладка"
        RightBody = "• Откуда и куда перемещают груз.`r`n• Свободный путь перемещения.`r`n• Подготовленное место укладки, подкладки и опоры."
    }
    @{
        Code = "S031-P04"
        Title = "Подвал. Проверь схемы и таблицы"
        LeftTitle = "Схемы строповки"
        LeftBody = "• Используем схему под конкретный груз.`r`n• Проверяем точки зацепки и способ обвязки.`r`n• Самовольную замену схемы не допускаем."
        RightTitle = "Таблицы и документы"
        RightBody = "• Грузоподъемность крана и стропов.`r`n• Технологическая карта или ППР.`r`n• Ограничения по условиям работ и расстояниям."
    }
    @{
        Code = "S031-P05"
        Title = "Подвал. Осмотри приспособления и тару"
        LeftTitle = "Что готовим"
        LeftBody = "• Стропы, траверсы, захваты, тару и вспомогательную оснастку.`r`n• СИЗ и инструмент, необходимые до начала работ."
        RightTitle = "Что проверяем"
        RightBody = "• Бирки, клейма и грузоподъемность.`r`n• Износ, надрывы, деформации и трещины.`r`n• Все дефектные приспособления сразу выводим из работы."
    }
)

foreach ($spec in $footerSpecs) {
    $newSlide = $insertAfter.Duplicate().Item(1)
    $insertAfter = $newSlide
    Build-TwoPanelFooter $newSlide $spec.Title $spec.Code $spec.LeftTitle $spec.LeftBody $spec.RightTitle $spec.RightBody
    $newSlides[$spec.Code] = $newSlide
}

# Refresh slide references after insertions.
$slideP01 = Find-CodeSlide $presentation "S031-P01"
$slideP02 = Find-CodeSlide $presentation "S031-P02"
$slideP03 = Find-CodeSlide $presentation "S031-P03"
$slideP04 = Find-CodeSlide $presentation "S031-P04"
$slideP05 = Find-CodeSlide $presentation "S031-P05"

# Hook up card navigation on S031.
$targets = @($slideP01, $slideP02, $slideP03, $slideP04, $slideP05)
for ($i = 0; $i -lt $targets.Count; $i++) {
    $base = "Codex_S031_Card$($i+1)"
    foreach ($suffix in @("_panel","_strip","_text")) {
        $shape = $slideS031.Shapes.Item("$base$suffix")
        Set-SlideJump $shape $targets[$i]
    }
}

# Footer navigation.
Set-SlideJump (Try-GetShapeById $slideP01 11) $slideS031
Set-SlideJump (Try-GetShapeById $slideP01 12) $slideP02

Set-SlideJump (Try-GetShapeById $slideP02 11) $slideS031
Set-SlideJump (Try-GetShapeById $slideP02 12) $slideP03

Set-SlideJump (Try-GetShapeById $slideP03 11) $slideS031
Set-SlideJump (Try-GetShapeById $slideP03 12) $slideP04

Set-SlideJump (Try-GetShapeById $slideP04 11) $slideS031
Set-SlideJump (Try-GetShapeById $slideP04 12) $slideP05

Set-SlideJump (Try-GetShapeById $slideP05 11) $slideS031
Set-SlideJump (Try-GetShapeById $slideP05 12) $slideS032

$presentation.Save()

if (-not $wasAlreadyOpen) {
    $presentation.Close()
}

if ($createdApp) {
    $ppt.Quit()
}
