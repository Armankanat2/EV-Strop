$ErrorActionPreference = "Stop"

$srcPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-30_v56_ev-group-overlay.pptx"
$outPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-30_v57_ev-group-overlay-safe-polish.pptx"

Copy-Item -LiteralPath $srcPath -Destination $outPath -Force

$msoTrue = -1
$ppAlignCenter = 2
$ppAlignLeft = 1

# PowerPoint COM expects BGR integers.
$ColorBlue = 0xFF5700
$ColorDeepBlue = 0xB83800
$ColorWhite = 0xFFFFFF
$ColorGraphite = 0x111111
$ColorDarkGray = 0x63554B
$ColorLightGray = 0xEBE7E5
$ColorPaleBlue = 0xFFF0EA

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

function Set-TextStyle {
    param(
        $shape,
        [string]$fontName,
        [double]$size,
        [bool]$bold,
        [int]$rgb,
        [int]$align = 1
    )

    if ($shape.HasTextFrame -ne $msoTrue) { return }
    if ($shape.TextFrame.HasText -ne $msoTrue) { return }

    $shape.TextFrame.MarginLeft = 7
    $shape.TextFrame.MarginRight = 7
    $shape.TextFrame.MarginTop = 5
    $shape.TextFrame.MarginBottom = 5
    $shape.TextFrame.WordWrap = $msoTrue

    $range = $shape.TextFrame.TextRange
    $range.Font.Name = $fontName
    $range.Font.Size = $size
    $range.Font.Bold = $(if ($bold) { $msoTrue } else { 0 })
    Set-Rgb $range.Font.Color $rgb
    $range.ParagraphFormat.Alignment = $align
}

function Style-PrimaryButton {
    param($shape)
    Set-FillSolid $shape $ColorDeepBlue
    Hide-Line $shape
    Set-TextStyle $shape "Arial" 12 $true $ColorWhite $ppAlignCenter
}

function Style-SecondaryButton {
    param($shape)
    Set-FillSolid $shape $ColorWhite
    Set-Line $shape $ColorLightGray 1.0
    Set-TextStyle $shape "Arial" 12 $true $ColorDarkGray $ppAlignCenter
}

function Style-HeaderPanel {
    param($shape)
    Set-FillSolid $shape $ColorPaleBlue
    Set-Line $shape $ColorLightGray 0.8
}

function Style-ContentPanel {
    param($shape)
    Set-FillSolid $shape $ColorWhite
    Set-Line $shape $ColorLightGray 1.15
}

function Style-InfoCard {
    param($shape)
    Set-FillSolid $shape $ColorPaleBlue
    Hide-Line $shape
}

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = $msoTrue
$presentation = $pp.Presentations.Open($outPath, $false, $false, $false)

foreach ($slide in $presentation.Slides) {
    $codeShape = $null
    $textShapes = @()

    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
            $text = $shape.TextFrame.TextRange.Text.Trim()
            $textShapes += [PSCustomObject]@{ Shape = $shape; Text = $text }
            if ($text -match '^S\d{3}(-P\d{2})?(-PP\d{2})?$') {
                $codeShape = $shape
            }
        }
    }

    foreach ($shape in $slide.Shapes) {
        if ($shape.Type -ne 1) { continue }

        $text = ""
        if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
            $text = $shape.TextFrame.TextRange.Text.Trim()
        }

        if ($text -match '^S\d{3}(-P\d{2})?(-PP\d{2})?$') { continue }
        if ($text -match '^(Назад|НАЗАД|Назад к .+)$') {
            Style-SecondaryButton $shape
            continue
        }
        if ($text -match '^(Далее|ДАЛЕЕ|НАЧАТЬ КУРС|ОТКРЫТЬ .+)$') {
            Style-PrimaryButton $shape
            continue
        }

        if ($shape.Top -lt 40 -and $shape.Width -gt 800) {
            Set-FillSolid $shape $ColorWhite
            Hide-Line $shape
            continue
        }

        if ($shape.Top -gt 45 -and $shape.Top -lt 80 -and $shape.Width -gt 300 -and $shape.Height -gt 200) {
            Set-FillSolid $shape $ColorWhite
            Set-Line $shape $ColorDeepBlue 1.2
            continue
        }

        if ($shape.Top -gt 95 -and $shape.Top -lt 170 -and $shape.Height -lt 42 -and $shape.Width -gt 100) {
            Style-HeaderPanel $shape
            continue
        }

        if ($shape.Top -gt 90 -and $shape.Top -lt 470 -and $shape.Width -gt 150 -and $shape.Height -gt 28) {
            Style-ContentPanel $shape
            continue
        }

        if ($text.Length -gt 45 -and $shape.Top -gt 220 -and $shape.Top -lt 380 -and $shape.Width -gt 300 -and $shape.Height -gt 40) {
            Style-InfoCard $shape
            continue
        }
    }

    if ($codeShape) {
        Set-FillSolid $codeShape $ColorDeepBlue
        Hide-Line $codeShape
        Set-TextStyle $codeShape "Arial" 12.5 $true $ColorWhite $ppAlignCenter
    }

    foreach ($entry in $textShapes) {
        $shape = $entry.Shape
        $text = $entry.Text

        if ($codeShape -and $shape.Id -eq $codeShape.Id) { continue }

        if ($text -match '^(Назад|НАЗАД|Назад к .+)$') {
            Set-TextStyle $shape "Arial" 12 $true $ColorDarkGray $ppAlignCenter
            continue
        }

        if ($text -match '^(Далее|ДАЛЕЕ|НАЧАТЬ КУРС|ОТКРЫТЬ .+)$') {
            Set-TextStyle $shape "Arial" 12 $true $ColorWhite $ppAlignCenter
            continue
        }

        if ($text -match '^Тема \d' -or $text -match '^Подвал от S\d{3}$') {
            Set-TextStyle $shape "Arial" 11 $false $ColorDarkGray $ppAlignLeft
            continue
        }

        if ($shape.Top -lt 45) {
            if ($shape.Width -gt 240) {
                Set-TextStyle $shape "Arial" 11 $true $ColorDarkGray $ppAlignLeft
            }
            continue
        }

        if ($shape.Top -lt 110 -and $shape.Width -gt 220) {
            Set-TextStyle $shape "Arial" 22 $true $ColorGraphite $ppAlignLeft
            continue
        }

        if ($shape.Top -gt 120 -and $shape.Top -lt 185 -and $shape.Width -lt 380) {
            Set-TextStyle $shape "Arial" 13 $true $ColorDeepBlue $ppAlignLeft
            continue
        }

        if ($shape.Top -gt 175 -and $shape.Top -lt 470) {
            if ($text -match '^(Нужно|Плейсхолдер)') {
                Set-TextStyle $shape "Arial" 12 $false $ColorDarkGray $ppAlignLeft
            }
            elseif ($text.Length -gt 110) {
                Set-TextStyle $shape "Arial" 12.5 $false $ColorGraphite $ppAlignLeft
            }
            else {
                Set-TextStyle $shape "Arial" 13 $false $ColorGraphite $ppAlignLeft
            }
        }
    }

    if ($codeShape -and $codeShape.TextFrame.TextRange.Text.Trim() -eq "S001") {
        foreach ($entry in $textShapes) {
            $shape = $entry.Shape
            if ($shape.Top -gt 125 -and $shape.Top -lt 190 -and $shape.Width -lt 360) {
                Set-TextStyle $shape "Arial" 13 $true $ColorDeepBlue $ppAlignLeft
            }
            elseif ($shape.Top -gt 190 -and $shape.Top -lt 240 -and $shape.Width -gt 300) {
                Set-TextStyle $shape "Arial" 15 $false $ColorGraphite $ppAlignLeft
            }
        }
    }
}

$presentation.Save()
$presentation.Close()
$pp.Quit()

Write-Output $outPath
