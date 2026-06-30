$ErrorActionPreference = "Stop"

$presentationPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S029_theme22_preview_2026-06-30_v13_ev-group-techstyle-premium-subslides.pptx"

$msoTrue = -1
$ppMouseClick = 1
$ppActionHyperlink = 7
$ppAlignLeft = 1

# PowerPoint COM expects BGR integers.
$ColorWhite = 0xFFFFFF
$ColorGraphite = 0x111111
$ColorDarkGray = 0x63554B
$ColorLightGray = 0xEBE7E5

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

function Get-ShapeById {
    param($slide, [int]$id)

    foreach ($shape in $slide.Shapes) {
        if ($shape.Id -eq $id) {
            return $shape
        }
    }

    throw "Shape with id '$id' not found on slide."
}

function Set-Rgb {
    param($holder, [int]$rgb)
    $holder.RGB = $rgb
}

function Set-TextFrameStyle {
    param(
        $shape,
        [double]$size,
        [bool]$bold = $false,
        [int]$rgb = $ColorGraphite,
        [int]$align = $ppAlignLeft
    )

    if ($shape.HasTextFrame -ne $msoTrue) { return }
    if ($shape.TextFrame.HasText -ne $msoTrue) { return }

    $shape.TextFrame.WordWrap = $msoTrue
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 6
    $shape.TextFrame.MarginBottom = 6

    $range = $shape.TextFrame.TextRange
    $range.Font.Name = "Arial"
    $range.Font.Size = $size
    $range.Font.Bold = $(if ($bold) { $msoTrue } else { 0 })
    Set-Rgb $range.Font.Color $rgb
    $range.ParagraphFormat.Alignment = $align
}

function Add-BodyTextbox {
    param(
        $slide,
        [double]$left,
        [double]$top,
        [double]$width,
        [double]$height,
        [string[]]$lines,
        [double]$size = 10.5
    )

    $textbox = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
    $textbox.TextFrame.WordWrap = $msoTrue
    $textbox.TextFrame.MarginLeft = 8
    $textbox.TextFrame.MarginRight = 8
    $textbox.TextFrame.MarginTop = 6
    $textbox.TextFrame.MarginBottom = 6
    $textbox.Fill.Visible = 0
    $textbox.Line.Visible = 0

    $text = [string]::Join("`r`n", $lines)
    $textbox.TextFrame.TextRange.Text = $text
    Set-TextFrameStyle $textbox $size $false $ColorGraphite $ppAlignLeft
    return $textbox
}

function Ensure-ActionToSlide {
    param($shape, $targetSlide)

    $shape.ActionSettings($ppMouseClick).Action = $ppActionHyperlink
    $shape.ActionSettings($ppMouseClick).Hyperlink.Address = ""
    $shape.ActionSettings($ppMouseClick).Hyperlink.SubAddress = $targetSlide.SlideID.ToString()
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

$presentation.Save()

$slideS031 = Find-CodeSlide $presentation "S031"
$slideS031P01 = Find-CodeSlide $presentation "S031-P01"

# Save current S031 edits and ensure the left panel can open the subslide.
foreach ($shapeId in @(6, 7, 8)) {
    try {
        Ensure-ActionToSlide (Get-ShapeById $slideS031 $shapeId) $slideS031P01
    }
    catch {
    }
}

# Update S031-P01 content.
(Get-ShapeById $slideS031P01 7).TextFrame.TextRange.Text = "Инструкция ""Уточни задание перед началом работ"""
Set-TextFrameStyle (Get-ShapeById $slideS031P01 7) 12.5 $true $ColorGraphite

$leftBodyLines = @(
    "Инструкция требует от работника получить четкие указания от лица, ответственного за безопасное производство работ (обычно это мастер или прораб).",
    "",
    "Что входит в уточнение задания",
    "• Место проведения работ: точная зона, где будут перемещать грузы.",
    "• Тип и масса груза: стропальщик должен знать точный вес и характер груза.",
    "• Схема строповки: способ обвязки или зацепки конкретного груза.",
    "• Выбор приспособлений: типы и грузоподъемность необходимых строп, траверс или захватов.",
    "• Маршрут перемещения: траектория движения груза и место его конечной укладки.",
    "• Опасные зоны: наличие вблизи ЛЭП, котлованов, заборов или других препятствий.",
    "• Знаки сигнализации: согласование порядка обмена сигналами с машинистом крана."
)

(Get-ShapeById $slideS031P01 8).TextFrame.TextRange.Text = [string]::Join("`r`n", $leftBodyLines)
Set-TextFrameStyle (Get-ShapeById $slideS031P01 8) 10.3 $false $ColorGraphite

(Get-ShapeById $slideS031P01 21).TextFrame.TextRange.Text = "Порядок действий стропальщика перед стартом"
Set-TextFrameStyle (Get-ShapeById $slideS031P01 21) 12.5 $true $ColorGraphite

# Remove the previous visual placeholder and caption from the right panel.
foreach ($shapeId in @(10, 22)) {
    try {
        (Get-ShapeById $slideS031P01 $shapeId).Delete()
    }
    catch {
    }
}

$rightBodyLines = @(
    "1. Получить инструктаж: выслушать задание от ответственного лица.",
    "2. Проверить документацию: ознакомиться с технологической картой или схемой строповки.",
    "3. Подобрать СИЗ: надеть каску, жилет, защитные перчатки и спецобувь.",
    "4. Осмотреть инструмент: проверить бирки и клейма на стропах, убедиться в отсутствии дефектов.",
    "5. Оценить площадку: проверить освещенность и отсутствие посторонних людей в зоне работы.",
    "",
    "Если схема строповки неизвестна или масса груза превышает грузоподъемность крана, стропальщик обязан отказаться от выполнения работ до устранения этих проблем."
)

$rightBody = Add-BodyTextbox $slideS031P01 406.8 202.0 458.0 214.0 $rightBodyLines 10.5

$presentation.Save()

if (-not $wasAlreadyOpen) {
    $presentation.Close()
}

if ($createdApp) {
    $ppt.Quit()
}
