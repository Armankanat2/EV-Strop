$ErrorActionPreference = "Stop"

$presentationPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S029_theme22_preview_2026-06-30_v13_ev-group-techstyle-premium-subslides.pptx"

$msoTrue = -1
$msoShapeRectangle = 1
$msoSendToBack = 1
$ppAlignLeft = 1

# PowerPoint COM expects BGR integers.
$ColorWhite = 0xFFFFFF
$ColorGraphite = 0x111111
$ColorDarkGray = 0x63554B
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

function Apply-BaseTextStyle {
    param($shape, [double]$size)

    $shape.TextFrame.WordWrap = $msoTrue
    $shape.TextFrame.MarginLeft = 8
    $shape.TextFrame.MarginRight = 8
    $shape.TextFrame.MarginTop = 6
    $shape.TextFrame.MarginBottom = 6

    $range = $shape.TextFrame.TextRange
    $range.Font.Name = "Arial"
    $range.Font.Size = $size
    $range.Font.Bold = 0
    Set-Rgb $range.Font.Color $ColorGraphite
    $range.ParagraphFormat.Alignment = $ppAlignLeft
}

function Highlight-Substring {
    param(
        $textRange,
        [string]$substring,
        [bool]$bold = $true,
        [int]$rgb = $ColorDeepBlue
    )

    $start = 0
    while ($true) {
        $index = $textRange.Text.IndexOf($substring, $start, [System.StringComparison]::Ordinal)
        if ($index -lt 0) { break }

        $chars = $textRange.Characters($index + 1, $substring.Length)
        $chars.Font.Bold = $(if ($bold) { $msoTrue } else { 0 })
        Set-Rgb $chars.Font.Color $rgb
        $start = $index + $substring.Length
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

$slide = Find-CodeSlide $presentation "S031-P01"

# Rebuild the content zone as a single large text panel.
$panel = Get-ShapeById $slide 6
$panel.Left = 53.3
$panel.Top = 144.0
$panel.Width = 825.8
$panel.Height = 308.2
Set-FillSolid $panel $ColorWhite
Set-Line $panel $ColorLightGray 1.15

$topStrip = Get-ShapeById $slide 14
$topStrip.Left = 53.3
$topStrip.Top = 144.0
$topStrip.Width = 825.8
$topStrip.Height = 2.9
Set-FillSolid $topStrip $ColorDeepBlue
Hide-Line $topStrip

foreach ($id in @(7, 8, 9, 17, 21, 23, 24)) {
    try {
        (Get-ShapeById $slide $id).Delete()
    }
    catch {
    }
}

$textBox = $slide.Shapes.AddTextbox(1, 71.0, 183.0, 790.0, 236.0)
$textBox.Fill.Visible = 0
$textBox.Line.Visible = 0

$leftHeader = 'Инструкция "Уточни задание перед началом работ"'
$leftBody = @(
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

$rightHeader = "Порядок действий стропальщика перед стартом"
$rightBody = @(
    "1. Получить инструктаж: выслушать задание от ответственного лица.",
    "2. Проверить документацию: ознакомиться с технологической картой или схемой строповки.",
    "3. Подобрать СИЗ: надеть каску, жилет, защитные перчатки и спецобувь.",
    "4. Осмотреть инструмент: проверить бирки и клейма на стропах, убедиться в отсутствии дефектов.",
    "5. Оценить площадку: проверить освещенность и отсутствие посторонних людей в зоне работы.",
    "",
    "Если схема строповки неизвестна или масса груза превышает грузоподъемность крана, стропальщик обязан отказаться от выполнения работ до устранения этих проблем."
)

$fullText = @(
    $leftHeader
    ""
    $leftBody
    ""
    $rightHeader
    ""
    $rightBody
) -join "`r`n"

$textBox.TextFrame.TextRange.Text = $fullText
Apply-BaseTextStyle $textBox 10.2

try {
    $textBox.TextFrame2.Column.Number = 2
    $textBox.TextFrame2.Column.Spacing = 18
}
catch {
}

Highlight-Substring $textBox.TextFrame.TextRange $leftHeader $true $ColorDeepBlue
Highlight-Substring $textBox.TextFrame.TextRange "Что входит в уточнение задания" $true $ColorDeepBlue
Highlight-Substring $textBox.TextFrame.TextRange $rightHeader $true $ColorDeepBlue
Highlight-Substring $textBox.TextFrame.TextRange "Если схема строповки неизвестна или масса груза превышает грузоподъемность крана, стропальщик обязан отказаться от выполнения работ до устранения этих проблем." $true $ColorDeepBlue

foreach ($phrase in @(
    "Место проведения работ:",
    "Тип и масса груза:",
    "Схема строповки:",
    "Выбор приспособлений:",
    "Маршрут перемещения:",
    "Опасные зоны:",
    "Знаки сигнализации:",
    "1. Получить инструктаж:",
    "2. Проверить документацию:",
    "3. Подобрать СИЗ:",
    "4. Осмотреть инструмент:",
    "5. Оценить площадку:"
)) {
    Highlight-Substring $textBox.TextFrame.TextRange $phrase $true $ColorGraphite
}

$presentation.Save()

if (-not $wasAlreadyOpen) {
    $presentation.Close()
}

if ($createdApp) {
    $ppt.Quit()
}
