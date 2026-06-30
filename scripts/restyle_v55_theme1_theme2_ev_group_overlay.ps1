$ErrorActionPreference = "Stop"

$srcPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-29_v55.pptx"
$outPath = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-30_v56_ev-group-overlay.pptx"

Copy-Item -LiteralPath $srcPath -Destination $outPath -Force

$msoTrue = -1
$msoShapeRectangle = 1
$ppAlignCenter = 2

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

    $shape.TextFrame.MarginLeft = 6
    $shape.TextFrame.MarginRight = 6
    $shape.TextFrame.MarginTop = 4
    $shape.TextFrame.MarginBottom = 4
    $shape.TextFrame.WordWrap = $msoTrue

    $range = $shape.TextFrame.TextRange
    $range.Font.Name = $fontName
    $range.Font.Size = $size
    $range.Font.Bold = $(if ($bold) { $msoTrue } else { 0 })
    Set-Rgb $range.Font.Color $rgb
    $range.ParagraphFormat.Alignment = $align
}

function Add-AccentLine {
    param($slide)
    $line = $slide.Shapes.AddShape($msoShapeRectangle, 40, 53, 82, 4)
    Set-FillSolid $line $ColorBlue
    Hide-Line $line
}

function Normalize-Panel {
    param($shape)
    Set-FillSolid $shape $ColorWhite
    Set-Line $shape $ColorLightGray 1.1
}

function Normalize-HeaderShape {
    param($shape)
    Set-FillSolid $shape $ColorPaleBlue
    Set-Line $shape $ColorLightGray 0.8
    Set-TextStyle $shape "Arial" 13 $true $ColorDeepBlue
}

function Normalize-BadgeShape {
    param($shape)
    Set-FillSolid $shape $ColorPaleBlue
    Set-Line $shape $ColorBlue 1.0
    Set-TextStyle $shape "Arial" 13 $true $ColorDeepBlue $ppAlignCenter
}

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = $msoTrue
$presentation = $pp.Presentations.Open($outPath, $false, $false, $false)

foreach ($slide in $presentation.Slides) {
    $codeShape = $null
    $textShapes = @()
    $navShapes = @()

    foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
            $text = $shape.TextFrame.TextRange.Text.Trim()
            $textShapes += [PSCustomObject]@{ Shape = $shape; Text = $text }

            if ($text -match '^S\d{3}(-P\d{2})?(-PP\d{2})?$') {
                $codeShape = $shape
            }
            elseif ($text -match '^(Назад|НАЗАД|Далее|ДАЛЕЕ|НАЧАТЬ КУРС|ОТКРЫТЬ .+)$') {
                $navShapes += [PSCustomObject]@{ Shape = $shape; Text = $text }
            }
        }
    }

    foreach ($shape in $slide.Shapes) {
        if ($shape.Type -eq 1) {
            $isCode = $false
            if ($codeShape -and $shape.Id -eq $codeShape.Id) { $isCode = $true }
            if ($isCode) { continue }

            $isNav = $false
            foreach ($n in $navShapes) {
                if ($shape.Id -eq $n.Shape.Id) { $isNav = $true; break }
            }
            if ($isNav) { continue }

            if ($shape.Top -lt 40 -and $shape.Width -gt 800) {
                Set-FillSolid $shape $ColorWhite
                Hide-Line $shape
            }
            elseif ($shape.Top -gt 90 -and $shape.Top -lt 460 -and $shape.Width -gt 150 -and $shape.Height -gt 28) {
                Normalize-Panel $shape
            }
            elseif ($shape.Top -gt 90 -and $shape.Top -lt 180 -and $shape.Height -lt 42 -and $shape.Width -gt 100) {
                Normalize-HeaderShape $shape
            }
            elseif ($shape.Top -gt 90 -and $shape.Top -lt 430 -and $shape.Width -lt 60 -and $shape.Height -lt 60) {
                $txt = ""
                if ($shape.HasTextFrame -eq $msoTrue -and $shape.TextFrame.HasText -eq $msoTrue) {
                    $txt = $shape.TextFrame.TextRange.Text.Trim()
                }
                if ($txt -match '^\d+$' -or $txt -match '^[А-ЯA-Z]$') {
                    Normalize-BadgeShape $shape
                }
            }
        }
    }

    if ($codeShape) {
        Set-FillSolid $codeShape $ColorDeepBlue
        Hide-Line $codeShape
        Set-TextStyle $codeShape "Arial" 13 $true $ColorWhite $ppAlignCenter
    }

    foreach ($entry in $textShapes) {
        $shape = $entry.Shape
        $text = $entry.Text

        if ($codeShape -and $shape.Id -eq $codeShape.Id) { continue }

        if ($text -match '^(Назад|НАЗАД)$') {
            Set-FillSolid $shape $ColorWhite
            Set-Line $shape $ColorLightGray 1.0
            Set-TextStyle $shape "Arial" 12 $true $ColorDarkGray $ppAlignCenter
            continue
        }

        if ($text -match '^(Далее|ДАЛЕЕ|НАЧАТЬ КУРС|ОТКРЫТЬ .+)$') {
            Set-FillSolid $shape $ColorBlue
            Hide-Line $shape
            Set-TextStyle $shape "Arial" 12 $true $ColorWhite $ppAlignCenter
            continue
        }

        if ($shape.Top -lt 45) {
            if ($shape.Height -gt 26) {
                Set-TextStyle $shape "Arial" 22 $true $ColorGraphite
            }
            else {
                Set-TextStyle $shape "Arial" 11 $true $ColorDarkGray
            }
            continue
        }

        if ($shape.Top -lt 110) {
            if ($shape.Height -gt 24) {
                Set-TextStyle $shape "Arial" 24 $true $ColorGraphite
            }
            else {
                Set-TextStyle $shape "Arial" 12 $false $ColorDarkGray
            }
            continue
        }

        if ($shape.Top -gt 90 -and $shape.Top -lt 180 -and $shape.Height -lt 42) {
            Set-TextStyle $shape "Arial" 14 $true $ColorGraphite
            continue
        }

        if ($shape.Top -gt 90 -and $shape.Top -lt 460) {
            Set-TextStyle $shape "Arial" 13 $false $ColorGraphite
        }
    }

    Add-AccentLine $slide
}

$presentation.Save()
$presentation.Close()
$pp.Quit()

Write-Output $outPath
