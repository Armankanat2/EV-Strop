$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "..\..\..\..\scripts\powerpoint_stable_tools.ps1")

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName System.Drawing

function RgbValue {
    param([int] $R, [int] $G, [int] $B)
    return $R + (256 * $G) + (65536 * $B)
}

function InchesToPoints {
    param([double] $Value)
    return [double]($Value * 72.0)
}

function Set-SlideJump {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)] $TargetSlide
    )

    $click = $Shape.ActionSettings(1)
    $click.Action = 7
    $click.Hyperlink.Address = ""
    $click.Hyperlink.SubAddress = "$($TargetSlide.SlideID),$($TargetSlide.SlideIndex), "
}

function Add-Box {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [int] $ShapeType,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [Parameter(Mandatory = $true)] [int] $FillColor,
        [Parameter(Mandatory = $true)] [int] $LineColor
    )

    $shape = $Slide.Shapes.AddShape(
        $ShapeType,
        (InchesToPoints $Left),
        (InchesToPoints $Top),
        (InchesToPoints $Width),
        (InchesToPoints $Height)
    )
    $shape.Fill.Solid()
    $shape.Fill.ForeColor.RGB = $FillColor
    $shape.Line.ForeColor.RGB = $LineColor
    return $shape
}

function Add-TextBox {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height
    )

    $shape = $Slide.Shapes.AddTextbox(
        1,
        (InchesToPoints $Left),
        (InchesToPoints $Top),
        (InchesToPoints $Width),
        (InchesToPoints $Height)
    )
    $shape.TextFrame.WordWrap = -1
    return $shape
}

function Set-TextShape {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)] [string] $Text,
        [int] $FontSize = 18,
        [bool] $Bold = $false,
        [int] $Color = 0,
        [int] $Align = 1
    )

    $Shape.TextFrame.TextRange.Text = $Text
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    if ($Color -ne 0) {
        $Shape.TextFrame.TextRange.Font.Color.RGB = $Color
    }
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
}

function Set-Lines {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)] [string[]] $Lines,
        [int] $FontSize = 18,
        [int] $Color = 0,
        [bool] $Bold = $false,
        [int] $Align = 1
    )

    $Shape.TextFrame.TextRange.Text = ($Lines -join "`r")
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    if ($Color -ne 0) {
        $Shape.TextFrame.TextRange.Font.Color.RGB = $Color
    }
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = $Align
}

function Set-TermDefinitionBlock {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)] [object[]] $Items,
        [double] $FontSize = 13,
        [int] $TextColor = 0,
        [int] $TermColor = 0
    )

    if ($TextColor -eq 0) {
        $TextColor = $script:TEXT
    }
    if ($TermColor -eq 0) {
        $TermColor = $script:BLUE
    }

    $lines = @()
    foreach ($item in $Items) {
        $lines += "$($item.Term) — $($item.Definition)"
    }

    $Shape.TextFrame.MarginLeft = 0
    $Shape.TextFrame.MarginRight = 0
    $Shape.TextFrame.MarginTop = 0
    $Shape.TextFrame.MarginBottom = 0
    $Shape.TextFrame.WordWrap = -1
    $Shape.TextFrame.VerticalAnchor = 1
    $Shape.TextFrame.TextRange.Text = ($lines -join "`r")
    $Shape.TextFrame.TextRange.Font.Name = "Arial"
    $Shape.TextFrame.TextRange.Font.Size = $FontSize
    $Shape.TextFrame.TextRange.Font.Bold = 0
    $Shape.TextFrame.TextRange.Font.Color.RGB = $TextColor
    $Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $Shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.0
    $Shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 2

    $position = 1
    foreach ($item in $Items) {
        $termLength = $item.Term.Length
        $definitionLength = $item.Definition.Length
        $termRange = $Shape.TextFrame.TextRange.Characters($position, $termLength)
        $termRange.Font.Bold = -1
        $termRange.Font.Color.RGB = $TermColor
        $position += $termLength + 3 + $definitionLength + 1
    }
}

function Add-Button {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [Parameter(Mandatory = $true)] [string] $Text,
        [Parameter(Mandatory = $true)] [int] $FillColor,
        [int] $FontSize = 18
    )

    $shape = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $FillColor -LineColor $FillColor
    Set-TextShape -Shape $shape -Text $Text -FontSize $FontSize -Bold $true -Color $script:WHITE -Align 2
    return $shape
}

function Add-Hotspot {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height
    )

    $shape = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRectangle) `
        -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $script:WHITE -LineColor $script:WHITE
    $shape.Fill.Transparency = 1.0
    $shape.Line.Transparency = 1.0
    return $shape
}

function Add-Band {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] [string] $Title,
        [string] $Subtitle = ""
    )

    $band = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left 0.45 -Top 0.42 -Width 12.12 -Height 0.88 -FillColor $script:BROWN -LineColor $script:BROWN
    $titleBox = Add-TextBox -Slide $Slide -Left 2.55 -Top 0.56 -Width 7.35 -Height 0.38
    Set-TextShape -Shape $titleBox -Text $Title -FontSize 27 -Bold $true -Color $script:WHITE

    if ($Subtitle) {
        $subBox = Add-TextBox -Slide $Slide -Left 2.56 -Top 0.98 -Width 7.15 -Height 0.24
        Set-TextShape -Shape $subBox -Text $Subtitle -FontSize 12 -Color $script:WHITE
    }

    $tag = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left 0.68 -Top 0.60 -Width 1.78 -Height 0.48 -FillColor $script:CARD -LineColor $script:CARD
    Set-TextShape -Shape $tag -Text $Code -FontSize 18 -Bold $true -Color $script:BROWN -Align 2
    return $band
}

function New-ThemeSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] [string] $Title,
        [string] $Subtitle = ""
    )

    $slide = $Presentation.Slides.Add($Presentation.Slides.Count + 1, 12)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = $script:BG
    Add-Band -Slide $slide -Code $Code -Title $Title -Subtitle $Subtitle | Out-Null
    return $slide
}

function Add-Panel {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [Parameter(Mandatory = $true)] [string] $Title,
        [int] $AccentColor = 0,
        [double] $HeaderWidth = 2.36,
        [int] $HeaderFontSize = 14
    )

    if ($AccentColor -eq 0) {
        $AccentColor = $script:BROWN
    }

    $panel = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $script:CARD -LineColor $script:LINE
    $head = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left ($Left + 0.16) -Top ($Top + 0.16) -Width $HeaderWidth -Height 0.42 -FillColor $AccentColor -LineColor $AccentColor
    Set-TextShape -Shape $head -Text $Title -FontSize $HeaderFontSize -Bold $true -Color $script:WHITE -Align 1
    return $panel
}

function Add-BulletPanel {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [string[]] $Lines
    )

    Add-Panel -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -Title $Title | Out-Null
    $box = Add-TextBox -Slide $Slide -Left ($Left + 0.38) -Top ($Top + 0.84) -Width ($Width - 0.70) -Height ($Height - 1.10)
    $bullets = @()
    foreach ($line in $Lines) {
        $bullets += "• $line"
    }
    Set-Lines -Shape $box -Lines $bullets -FontSize 19 -Color $script:TEXT
    return $box
}

function Add-ImagePanel {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [string] $ImagePath,
        [string] $Caption = ""
    )

    Add-Panel -Slide $Slide -Left $Left -Top $Top -Width $Width -Height $Height -Title $Title | Out-Null
    $imageWidth = $Width - 0.60
    $imageHeight = $imageWidth * 9.0 / 16.0
    $topOffset = $Top + 0.82
    $Slide.Shapes.AddPicture(
        $ImagePath,
        0,
        -1,
        (InchesToPoints ($Left + 0.30)),
        (InchesToPoints $topOffset),
        (InchesToPoints $imageWidth),
        (InchesToPoints $imageHeight)
    ) | Out-Null

    if ($Caption) {
        $cap = Add-TextBox -Slide $Slide -Left ($Left + 0.32) -Top ($topOffset + $imageHeight + 0.18) -Width ($Width - 0.64) -Height 0.56
        Set-TextShape -Shape $cap -Text $Caption -FontSize 13 -Color $script:MUTED
    }
}

function Add-FitPicture {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [string] $ImagePath,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height
    )

    $boxLeft = InchesToPoints $Left
    $boxTop = InchesToPoints $Top
    $boxWidth = InchesToPoints $Width
    $boxHeight = InchesToPoints $Height
    $image = [System.Drawing.Image]::FromFile($ImagePath)

    try {
        $imageRatio = [double]$image.Width / [double]$image.Height
        $boxRatio = $boxWidth / $boxHeight

        if ($imageRatio -ge $boxRatio) {
            $drawWidth = $boxWidth
            $drawHeight = $boxWidth / $imageRatio
            $drawLeft = $boxLeft
            $drawTop = $boxTop + (($boxHeight - $drawHeight) / 2.0)
        }
        else {
            $drawHeight = $boxHeight
            $drawWidth = $boxHeight * $imageRatio
            $drawTop = $boxTop
            $drawLeft = $boxLeft + (($boxWidth - $drawWidth) / 2.0)
        }

        return $Slide.Shapes.AddPicture($ImagePath, 0, -1, $drawLeft, $drawTop, $drawWidth, $drawHeight)
    }
    finally {
        $image.Dispose()
    }
}

function Add-ImageLabelTile {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [string] $ImagePath = "",
        [Parameter(Mandatory = $true)] [string] $Label,
        [int] $AccentColor = 0
    )

    if ($AccentColor -eq 0) {
        $AccentColor = $script:BROWN
    }

    $tile = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $script:WHITE -LineColor $script:LINE
    $tile.Line.Weight = 1.2

    $imageBox = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left ($Left + 0.16) -Top ($Top + 0.14) -Width 1.74 -Height ($Height - 0.28) -FillColor $script:CARD -LineColor $script:LINE
    $imageBox.Line.Visible = 0
    if ($ImagePath -and (Test-Path -LiteralPath $ImagePath)) {
        Add-FitPicture -Slide $Slide -ImagePath $ImagePath -Left ($Left + 0.22) -Top ($Top + 0.18) -Width 1.62 -Height ($Height - 0.36) | Out-Null
    }

    $accent = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left ($Left + 2.00) -Top ($Top + 0.22) -Width 0.10 -Height ($Height - 0.44) -FillColor $AccentColor -LineColor $AccentColor
    $accent.Line.Visible = 0

    $labelBox = Add-TextBox -Slide $Slide -Left ($Left + 2.24) -Top ($Top + 0.22) -Width ($Width - 2.50) -Height ($Height - 0.22)
    $labelBox.TextFrame.VerticalAnchor = 3
    Set-TextShape -Shape $labelBox -Text $Label -FontSize 17 -Bold $true -Color $script:TEXT
    return $tile
}

function Add-CraneLabelItem {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [string] $ImagePath,
        [Parameter(Mandatory = $true)] [string] $Label,
        [Parameter(Mandatory = $true)] [double] $ImageLeft,
        [Parameter(Mandatory = $true)] [double] $ImageTop,
        [Parameter(Mandatory = $true)] [double] $ImageWidth,
        [Parameter(Mandatory = $true)] [double] $ImageHeight,
        [Parameter(Mandatory = $true)] [double] $LabelLeft,
        [Parameter(Mandatory = $true)] [double] $LabelTop,
        [double] $LabelWidth = 2.92,
        [double] $LabelHeight = 0.41
    )

    Add-FitPicture -Slide $Slide -ImagePath $ImagePath -Left $ImageLeft -Top $ImageTop -Width $ImageWidth -Height $ImageHeight | Out-Null

    $labelBox = Add-TextBox -Slide $Slide -Left $LabelLeft -Top $LabelTop -Width $LabelWidth -Height $LabelHeight
    $labelBox.TextFrame.VerticalAnchor = 3
    Set-TextShape -Shape $labelBox -Text $Label -FontSize 18 -Bold $true -Color $script:TEXT
    return $labelBox
}

function Add-RouteCard {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [double] $Height,
        [Parameter(Mandatory = $true)] [string] $Step,
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [string] $Body,
        [Parameter(Mandatory = $true)] [int] $Accent
    )

    $card = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left $Left -Top $Top -Width $Width -Height $Height -FillColor $script:CARD -LineColor $Accent
    $chip = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeOval) `
        -Left ($Left + 0.18) -Top ($Top + 0.18) -Width 0.48 -Height 0.48 -FillColor $Accent -LineColor $Accent
    Set-TextShape -Shape $chip -Text $Step -FontSize 14 -Bold $true -Color $script:WHITE -Align 2
    $titleBox = Add-TextBox -Slide $Slide -Left ($Left + 0.82) -Top ($Top + 0.18) -Width ($Width - 1.05) -Height 0.32
    Set-TextShape -Shape $titleBox -Text $Title -FontSize 16 -Bold $true -Color $Accent
    $bodyBox = Add-TextBox -Slide $Slide -Left ($Left + 0.18) -Top ($Top + 0.74) -Width ($Width - 0.36) -Height ($Height - 0.92)
    Set-Lines -Shape $bodyBox -Lines @($Body) -FontSize 13 -Color $script:TEXT
    return $card
}

function Add-Nav {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [string] $PrevText,
        [string] $NextText
    )

    $prev = $null
    $next = $null
    if ($PrevText) {
        $prev = Add-Button -Slide $Slide -Left 0.70 -Top 6.62 -Width 2.28 -Height 0.50 -Text $PrevText -FillColor $script:STEEL -FontSize 16
    }
    if ($NextText) {
        $next = Add-Button -Slide $Slide -Left 10.30 -Top 6.62 -Width 2.28 -Height 0.50 -Text $NextText -FillColor $script:BROWN -FontSize 16
    }
    return @($prev, $next)
}

function Add-StepRow {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Left,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [double] $Width,
        [Parameter(Mandatory = $true)] [string] $IndexText,
        [Parameter(Mandatory = $true)] [string] $Text
    )

    $row = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left $Left -Top $Top -Width $Width -Height 0.62 -FillColor $script:WHITE -LineColor $script:LINE
    $dot = Add-Box -Slide $Slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeOval) `
        -Left ($Left + 0.16) -Top ($Top + 0.09) -Width 0.44 -Height 0.44 -FillColor $script:BROWN -LineColor $script:BROWN
    Set-TextShape -Shape $dot -Text $IndexText -FontSize 14 -Bold $true -Color $script:WHITE -Align 2
    $label = Add-TextBox -Slide $Slide -Left ($Left + 0.80) -Top ($Top + 0.16) -Width ($Width - 1.00) -Height 0.28
    Set-TextShape -Shape $label -Text $Text -FontSize 16 -Bold $true -Color $script:TEXT
    return $row
}

function Add-ArrowCallout {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $TextLeft,
        [Parameter(Mandatory = $true)] [double] $TextTop,
        [Parameter(Mandatory = $true)] [double] $TextWidth,
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [double] $TargetX,
        [Parameter(Mandatory = $true)] [double] $TargetY,
        [double] $LineWidth = 2.55
    )

    $titleBox = Add-TextBox -Slide $Slide -Left $TextLeft -Top $TextTop -Width $TextWidth -Height 0.24
    Set-TextShape -Shape $titleBox -Text $Title -FontSize 15 -Bold $true -Color $script:TEXT

    $underlineY = $TextTop + 0.30
    $lineLeft = $TextLeft
    $lineRight = $TextLeft + $LineWidth

    $underline = $Slide.Shapes.AddLine(
        (InchesToPoints $lineLeft),
        (InchesToPoints $underlineY),
        (InchesToPoints $lineRight),
        (InchesToPoints $underlineY)
    )
    $underline.Line.ForeColor.RGB = $script:OCHRE
    $underline.Line.Weight = 1.2

    $leader = $Slide.Shapes.AddLine(
        (InchesToPoints $lineLeft),
        (InchesToPoints $underlineY),
        (InchesToPoints $TargetX),
        (InchesToPoints $TargetY)
    )
    $leader.Line.ForeColor.RGB = $script:OCHRE
    $leader.Line.Weight = 1.15
    $leader.Line.EndArrowheadStyle = [Microsoft.Office.Core.MsoArrowheadStyle]::msoArrowheadTriangle
    $leader.Line.EndArrowheadLength = [Microsoft.Office.Core.MsoArrowheadLength]::msoArrowheadLengthMedium
    $leader.Line.EndArrowheadWidth = [Microsoft.Office.Core.MsoArrowheadWidth]::msoArrowheadNarrow

    return $titleBox
}

function New-DetailSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [string] $Lead,
        [Parameter(Mandatory = $true)] $BackTarget
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title $Title -Subtitle "Подвал к слайду S008"
    $leadBox = Add-TextBox -Slide $slide -Left 0.82 -Top 1.70 -Width 11.20 -Height 0.52
    Set-TextShape -Shape $leadBox -Text $Lead -FontSize 23 -Bold $true -Color $script:TEXT

    Add-BulletPanel -Slide $slide -Left 0.82 -Top 2.44 -Width 5.58 -Height 3.24 -Title "Назначение подвала" -Lines @(
        "Здесь раскрываем тему подробнее, без перегруза основного линейного слайда.",
        "Содержимое можно достраивать по ходу сборки темы 2.",
        "Структура уже готова для дальнейшего наполнения."
    ) | Out-Null

    Add-BulletPanel -Slide $slide -Left 6.82 -Top 2.44 -Width 5.46 -Height 3.24 -Title "Что сюда можно добавить" -Lines @(
        "Краткие пояснения и классификацию.",
        "Опорные схемы, таблицы или фото.",
        "Практические акценты для стропальщика."
    ) | Out-Null

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.58 -Height 0.50 -Text "Вернуться к S008" -FillColor $script:STEEL -FontSize 16
    Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    return $slide
}

function New-CraneTypesSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] $BackTarget,
        [Parameter(Mandatory = $true)] [hashtable] $ImageMap,
        $DetailTarget = $null
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title "Типы кранов" -Subtitle "Подвал к слайду S008"
    Add-Panel -Slide $slide -Left 0.82 -Top 1.64 -Width 11.46 -Height 4.90 -Title "Типы кранов" -AccentColor $script:BROWN | Out-Null

    if ($DetailTarget) {
        $detailBtn = Add-Button -Slide $slide -Left 7.92 -Top 1.76 -Width 3.54 -Height 0.47 -Text "Для чего нужны - описание" -FillColor $script:BLUE -FontSize 14
        Set-SlideJump -Shape $detailBtn -TargetSlide $DetailTarget
    }

    Add-CraneLabelItem -Slide $slide -ImagePath $ImageMap["tower"] -Label "Башенные краны" `
        -ImageLeft 1.28 -ImageTop 2.25 -ImageWidth 1.64 -ImageHeight 1.00 -LabelLeft 3.19 -LabelTop 2.38 | Out-Null
    Add-CraneLabelItem -Slide $slide -ImagePath $ImageMap["boom"] -Label "Стреловые краны" `
        -ImageLeft 6.56 -ImageTop 2.25 -ImageWidth 1.64 -ImageHeight 1.00 -LabelLeft 8.47 -LabelTop 2.38 | Out-Null
    Add-CraneLabelItem -Slide $slide -ImagePath $ImageMap["bridge"] -Label "Мостовые краны" `
        -ImageLeft 1.28 -ImageTop 3.42 -ImageWidth 1.64 -ImageHeight 1.00 -LabelLeft 3.19 -LabelTop 3.55 | Out-Null
    Add-CraneLabelItem -Slide $slide -ImagePath $ImageMap["gantry"] -Label "Козловые краны" `
        -ImageLeft 6.56 -ImageTop 3.42 -ImageWidth 1.64 -ImageHeight 1.00 -LabelLeft 8.47 -LabelTop 3.55 | Out-Null
    Add-CraneLabelItem -Slide $slide -ImagePath $ImageMap["manipulator"] -Label "Кран манипулятор" `
        -ImageLeft 1.28 -ImageTop 4.58 -ImageWidth 1.64 -ImageHeight 1.00 -LabelLeft 3.19 -LabelTop 4.71 | Out-Null
    Add-CraneLabelItem -Slide $slide -ImagePath $ImageMap["pipe"] -Label "Кран трубоукладчик" `
        -ImageLeft 6.56 -ImageTop 4.58 -ImageWidth 1.64 -ImageHeight 1.00 -LabelLeft 8.47 -LabelTop 4.71 | Out-Null

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.58 -Height 0.50 -Text "Вернуться к S008" -FillColor $script:STEEL -FontSize 16
    Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    return $slide
}

function New-CranePurposeSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        $BackTarget = $null
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title "Для чего нужны" -Subtitle "Поподвал к слайду S008-P01"
    Add-Panel -Slide $slide -Left 0.82 -Top 1.64 -Width 11.46 -Height 4.92 -Title "Для чего нужны" -AccentColor $script:BROWN | Out-Null

    $body = Add-TextBox -Slide $slide -Left 1.10 -Top 2.22 -Width 10.86 -Height 3.96
    Set-TextShape -Shape $body -Text (
        "Стреловые — для мобильных работ. Сами приезжают на объект. Нужны для разгрузки машин и мелких рассредоточенных строек.`r`r" +
        "Башенные — для многоэтажного строительства. Поднимают грузы на любую высоту и занимают минимум места на земле.`r`r" +
        "Мостовые — для заводов и цехов. Ездят под потолком, переносят тяжелые детали и не занимают место на полу.`r`r" +
        "Козловые — для открытых складов. Стоят на «ногах», разгружают ж/д вагоны, контейнеры и длинномерные грузы.`r`r" +
        "Портальные — для морских и речных портов. Разгружают корабли, пропуская под собой поезда и грузовики."
    ) -FontSize 16 -Color $script:TEXT

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.28 -Height 0.50 -Text "Назад" -FillColor $script:STEEL -FontSize 16
    if ($BackTarget) {
        Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    }
    return $slide
}

function New-CargoCategoryDetailSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] [string] $Title,
        [Parameter(Mandatory = $true)] [string] $ItemTitle,
        [Parameter(Mandatory = $true)] [string] $Definition,
        [Parameter(Mandatory = $true)] [string] $ImagePath,
        [Parameter(Mandatory = $true)] $BackTarget,
        [int] $AccentColor = 0,
        [string] $BodyText = "",
        [int] $BodyFontSize = 18
    )

    if ($AccentColor -eq 0) {
        $AccentColor = $script:BROWN
    }

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title $Title -Subtitle "Что влияет на устойчивость и выбор схемы"

    Add-Panel -Slide $slide -Left 0.82 -Top 1.66 -Width 4.26 -Height 4.80 -Title "Пример груза" -AccentColor $script:BROWN -HeaderWidth 2.20 | Out-Null
    $imageFrame = Add-Box -Slide $slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left 1.02 -Top 2.12 -Width 3.86 -Height 3.86 -FillColor $script:WHITE -LineColor $script:LINE
    $imageFrame.Line.Weight = 1.0
    Add-FitPicture -Slide $slide -ImagePath $ImagePath -Left 1.14 -Top 2.24 -Width 3.62 -Height 3.62 | Out-Null

    Add-Panel -Slide $slide -Left 5.30 -Top 1.66 -Width 6.98 -Height 4.80 -Title "Короткое пояснение" -AccentColor $AccentColor -HeaderWidth 3.18 | Out-Null
    $body = Add-TextBox -Slide $slide -Left 5.64 -Top 2.24 -Width 6.34 -Height 3.26
    if ([string]::IsNullOrWhiteSpace($BodyText)) {
        Set-TermDefinitionBlock -Shape $body -Items @(
            @{ Term = $ItemTitle; Definition = $Definition }
        ) -FontSize 18 -TextColor $script:TEXT -TermColor $AccentColor
    }
    else {
        Set-TextShape -Shape $body -Text $BodyText -FontSize $BodyFontSize -Color $script:TEXT
    }

    $note = Add-TextBox -Slide $slide -Left 5.64 -Top 5.18 -Width 6.18 -Height 0.64
    Set-Lines -Shape $note -Lines @(
        "До подъема контролируем точки зацепки, баланс и поведение груза."
    ) -FontSize 14 -Color $script:MUTED

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.28 -Height 0.50 -Text "Назад" -FillColor $script:STEEL -FontSize 16
    Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    return $slide
}

function New-CraneParametersSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] $BackTarget
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title "Основные параметры крана" -Subtitle "Подвал к слайду S008"
    Add-Panel -Slide $slide -Left 0.82 -Top 1.64 -Width 11.46 -Height 4.92 -Title "Основные параметры подъемных кранов" -AccentColor $script:BLUE | Out-Null

    $body = Add-TextBox -Slide $slide -Left 1.08 -Top 2.38 -Width 10.88 -Height 3.72
    Set-TermDefinitionBlock -Shape $body -FontSize 14 -TextColor $script:TEXT -TermColor $script:BLUE -Items @(
        @{ Term = "Грузоподъемность"; Definition = "предельный вес груза, который кран может поднять (уже включает массу строп, траверс и крюка)." },
        @{ Term = "Длина вылета"; Definition = "расстояние от шарнира (пяты) стрелы до оси головных блоков." },
        @{ Term = "Вылет крюка"; Definition = "расстояние по горизонтали от оси вращения крана до вертикальной оси крюка на ровной площадке." },
        @{ Term = "Грузовая характеристика"; Definition = "график или таблица, показывающая, какой вес кран может безопасно поднять на каждом вылете." },
        @{ Term = "Грузовой момент"; Definition = "произведение веса груза на вылет стрелы; главный показатель устойчивости крана против опрокидывания." },
        @{ Term = "Высота подъема"; Definition = "расстояние от земли (уровня стоянки) до крюка в его крайнем верхнем положении." },
        @{ Term = "Глубина опускания"; Definition = "расстояние от земли до крюка в его крайнем нижнем положении (в котловане или колодце)." },
        @{ Term = "Скорость перемещения груза"; Definition = "скорость вертикального подъема или опускания максимального для крана веса." },
        @{ Term = "Скорость посадки"; Definition = "минимальная скорость плавного опускания груза для его точной и безопасной укладки (монтажа)." },
        @{ Term = "Скорость перемещения"; Definition = "скорость горизонтального передвижения самого крана по площадке с грузом на крюке." },
        @{ Term = "Частота обращения"; Definition = "скорость вращения поворотной платформы крана со стрелой и грузом вокруг своей оси." }
    )

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.58 -Height 0.50 -Text "Вернуться к S008" -FillColor $script:STEEL -FontSize 16
    Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    return $slide
}

function Add-ConstructionRow {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [double] $Top,
        [Parameter(Mandatory = $true)] [string] $Term,
        [Parameter(Mandatory = $true)] [string] $Definition,
        [int] $TermColor = 0
    )

    if ($TermColor -eq 0) {
        $TermColor = $script:GREEN
    }

    $termBox = Add-TextBox -Slide $Slide -Left 1.08 -Top $Top -Width 2.40 -Height 0.34
    Set-TextShape -Shape $termBox -Text "• $Term" -FontSize 16 -Bold $true -Color $TermColor

    $defBox = Add-TextBox -Slide $Slide -Left 3.24 -Top $Top -Width 8.62 -Height 0.56
    Set-TextShape -Shape $defBox -Text $Definition -FontSize 15 -Color $script:TEXT
    return @($termBox, $defBox)
}

function New-CraneConstructionSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        [Parameter(Mandatory = $true)] $BackTarget,
        $HookTarget = $null,
        $HookAssemblyTarget = $null
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title "Конструктивные особенности" -Subtitle "Подвал к слайду S008"
    Add-Panel -Slide $slide -Left 0.82 -Top 1.64 -Width 11.46 -Height 4.92 -Title "Что нужно знать о конструкции крана" -AccentColor $script:GREEN | Out-Null

    Add-ConstructionRow -Slide $slide -Top 2.20 -Term "Опора" -Definition "колеса, рельсы или фундамент. Держит баланс, исключает опрокидывание." | Out-Null
    Add-ConstructionRow -Slide $slide -Top 2.64 -Term "Стрела (мост)" -Definition "металлоконструкция. Задает направление движения груза." | Out-Null
    Add-ConstructionRow -Slide $slide -Top 3.08 -Term "Тележка" -Definition "подвижная платформа на стреле или мосту. Перемещает груз вдоль них." | Out-Null
    Add-ConstructionRow -Slide $slide -Top 3.52 -Term "Лебедка" -Definition "барабан с мотором. Наматывает канат для подъема и опускания." | Out-Null
    Add-ConstructionRow -Slide $slide -Top 3.96 -Term "Тормоз" -Definition "блокиратор лебедки. Фиксирует канат, чтобы груз не упал под своим весом." | Out-Null

    $hookBtn = Add-Button -Slide $slide -Left 1.08 -Top 4.40 -Width 2.10 -Height 0.34 -Text "Крюк" -FillColor $script:BLUE -FontSize 14
    $hookDef = Add-TextBox -Slide $slide -Left 3.24 -Top 4.40 -Width 8.62 -Height 0.38
    Set-TextShape -Shape $hookDef -Text "точка зацепа. На него вешается грузозахватное приспособление." -FontSize 15 -Color $script:TEXT

    $hookAssemblyBtn = Add-Button -Slide $slide -Left 1.08 -Top 4.84 -Width 2.10 -Height 0.34 -Text "Крюковая подвеска" -FillColor $script:BLUE -FontSize 12
    $hookAssemblyDef = Add-TextBox -Slide $slide -Left 3.24 -Top 4.84 -Width 8.62 -Height 0.42
    Set-TextShape -Shape $hookAssemblyDef -Text "подвижный блок, удерживающий крюк и натягивающий канат." -FontSize 15 -Color $script:TEXT

    $summary = Add-TextBox -Slide $slide -Left 1.08 -Top 5.44 -Width 10.90 -Height 0.52
    Set-TextShape -Shape $summary -Text "Суть работы: Опора держит кран, стрела задает направление, тележка везет груз вдоль стрелы, лебедка мотает канат, тормоз держит вес, а крюк держит стропы." -FontSize 14 -Bold $true -Color $script:TEXT

    if ($HookTarget) {
        Set-SlideJump -Shape $hookBtn -TargetSlide $HookTarget
    }
    if ($HookAssemblyTarget) {
        Set-SlideJump -Shape $hookAssemblyBtn -TargetSlide $HookAssemblyTarget
    }

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.58 -Height 0.50 -Text "Вернуться к S008" -FillColor $script:STEEL -FontSize 16
    Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    return $slide
}

function New-HookAssemblySlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        $BackTarget = $null
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title "Крюк и крюковая подвеска" -Subtitle "Поподвал к слайду S008-P03"

    Add-Panel -Slide $slide -Left 0.82 -Top 1.64 -Width 5.44 -Height 4.66 -Title "Схема крюковой подвески" -AccentColor $script:BLUE -HeaderWidth 3.10 | Out-Null
    $schemeFrame = Add-Box -Slide $slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left 1.08 -Top 2.34 -Width 4.92 -Height 3.70 -FillColor $script:WHITE -LineColor $script:LINE
    $schemeFrame.Line.Weight = 1.0
    Add-FitPicture -Slide $slide -ImagePath $script:HookAssemblySchemeImagePath -Left 1.18 -Top 2.42 -Width 4.72 -Height 3.52 | Out-Null

    Add-Panel -Slide $slide -Left 6.48 -Top 1.64 -Width 5.80 -Height 4.66 -Title "Назначение крюковой подвески" -AccentColor $script:GREEN -HeaderWidth 3.54 | Out-Null
    $purposeText = Add-TextBox -Slide $slide -Left 6.84 -Top 2.72 -Width 5.06 -Height 3.18
    Set-TermDefinitionBlock -Shape $purposeText -Items @(
        @{ Term = "Соединение"; Definition = "связывает стальной канат крана со стропами груза." },
        @{ Term = "Выигрыш в силе"; Definition = "за счет встроенных блоков (полиспаста) снижает нагрузку на канат и лебедку при подъеме." },
        @{ Term = "Натяжение канатов"; Definition = "своей массой натягивает пустой канат, не давая ему запутаться на барабане." }
    ) -FontSize 16 -TextColor $script:TEXT -TermColor $script:GREEN

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.28 -Height 0.50 -Text "Назад" -FillColor $script:STEEL -FontSize 16
    if ($BackTarget) {
        Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    }
    return $slide
}

function New-HookDetailSlide {
    param(
        [Parameter(Mandatory = $true)] $Presentation,
        [Parameter(Mandatory = $true)] [string] $Code,
        $BackTarget = $null
    )

    $slide = New-ThemeSlide -Presentation $Presentation -Code $Code -Title "Крюк и крюковая подвеска" -Subtitle "Поподвал к слайду S008-P03"

    Add-Panel -Slide $slide -Left 0.82 -Top 1.64 -Width 5.44 -Height 4.66 -Title "Крюк/гак" -AccentColor $script:BLUE -HeaderWidth 1.72 | Out-Null
    $schemeFrame = Add-Box -Slide $slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left 1.08 -Top 2.12 -Width 4.92 -Height 3.92 -FillColor $script:WHITE -LineColor $script:LINE
    $schemeFrame.Line.Weight = 1.0
    Add-FitPicture -Slide $slide -ImagePath $script:HookSchemeImagePath -Left 1.18 -Top 2.18 -Width 4.72 -Height 3.78 | Out-Null

    Add-Panel -Slide $slide -Left 6.48 -Top 1.64 -Width 5.80 -Height 4.66 -Title "Крюк: что проверять при выбраковке" -AccentColor $script:RED -HeaderWidth 4.82 -HeaderFontSize 13 | Out-Null
    $rejectFrame = Add-Box -Slide $slide -ShapeType ([Microsoft.Office.Core.MsoAutoShapeType]::msoShapeRoundedRectangle) `
        -Left 6.76 -Top 2.18 -Width 1.82 -Height 3.10 -FillColor $script:WHITE -LineColor $script:LINE
    $rejectFrame.Line.Weight = 1.0
    Add-FitPicture -Slide $slide -ImagePath $script:HookRejectImagePath -Left 6.86 -Top 2.28 -Width 1.62 -Height 2.90 | Out-Null

    $rejectNotes = Add-TextBox -Slide $slide -Left 8.94 -Top 2.34 -Width 2.90 -Height 2.84
    $rejectNotes.TextFrame.MarginLeft = 0
    $rejectNotes.TextFrame.MarginRight = 0
    $rejectNotes.TextFrame.MarginTop = 0
    $rejectNotes.TextFrame.MarginBottom = 0
    Set-Lines -Shape $rejectNotes -Lines @(
        "1. Износ шейки крюка",
        "2. Отсутствие предохранительного устройства",
        "3. Износ более 10%",
        "4. Клюв крюка отогнут"
    ) -FontSize 16 -Color $script:TEXT -Bold $true

    $rejectNoteSmall = Add-TextBox -Slide $slide -Left 8.94 -Top 4.72 -Width 2.90 -Height 1.12
    $rejectNoteSmall.TextFrame.MarginLeft = 0
    $rejectNoteSmall.TextFrame.MarginRight = 0
    $rejectNoteSmall.TextFrame.MarginTop = 0
    $rejectNoteSmall.TextFrame.MarginBottom = 0
    Set-Lines -Shape $rejectNoteSmall -Lines @(
        "Запорная планка может отсутствовать при:",
        "Горячие цеха — при работе с расплавленным металлом и шлаком, так как замок плавится или деформируется от жары.",
        "Автоматические захваты — при дистанционной расстроповке, где замок технически мешает автоматическому сбросу строп."
    ) -FontSize 10 -Color $script:MUTED

    $backBtn = Add-Button -Slide $slide -Left 0.72 -Top 6.62 -Width 2.28 -Height 0.50 -Text "Назад" -FillColor $script:STEEL -FontSize 16
    if ($BackTarget) {
        Set-SlideJump -Shape $backBtn -TargetSlide $BackTarget
    }
    return $slide
}

$BG = RgbValue 247 243 235
$CARD = RgbValue 255 251 246
$WHITE = RgbValue 255 255 255
$BROWN = RgbValue 156 103 31
$TEXT = RgbValue 45 49 54
$MUTED = RgbValue 104 110 118
$LINE = RgbValue 220 207 188
$STEEL = RgbValue 96 113 136
$GREEN = RgbValue 42 143 82
$BLUE = RgbValue 72 116 157
$RED = RgbValue 212 74 62
$OCHRE = RgbValue 207 141 62

$workspaceRoot = (Resolve-Path (Join-Path $scriptDir "..\..\..\..")).Path
$livePreviewDir = Join-Path $workspaceRoot "presentations\ispring-course\module-01-stropovka-gruzov\live-preview"
$sourcePath = Join-Path $livePreviewDir "S001-S007_live_preview_working_2026-06-24_v32.pptx"
$outputPath = Join-Path $livePreviewDir "S001-S021_live_preview_working_2026-06-25_v38.pptx"

$assetRoot = Join-Path $workspaceRoot "assets\course-media\module-01-stropovka-gruzov\diagrams\error-analysis"
$imageMap = @{
    S011 = "S018-P03_tsentr-tyazhesti_diagram.png"
    S013 = "S018-P06_zatsepka-osnastka_checkpoints.png"
    S015 = "S018-P04_ugol-vetvey_nagruzka-diagram.png"
    S016 = "S018-P10_proverka-skhemy_steps.png"
    S017 = "S018-P07_stop-pri-opasnosti_poster.png"
}

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Base presentation not found: $sourcePath"
}

if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

Copy-Item -LiteralPath $sourcePath -Destination $outputPath -Force

$substDrive = $null
$pp = $null
$presentation = $null

try {
    $substDrive = Get-FreeSubstDrive
    & subst $substDrive $workspaceRoot | Out-Null
    $aliasRoot = $substDrive
    $outputAlias = Join-Path $aliasRoot "presentations\ispring-course\module-01-stropovka-gruzov\live-preview\S001-S021_live_preview_working_2026-06-25_v38.pptx"
    $craneImageRoot = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\cranes"
    $visualBankRoot = Join-Path $aliasRoot "assets\reference_visuals\visual-bank\images"
    $script:HookSchemeImagePath = Join-Path $visualBankRoot "VIS-0020_hook-safety-diagram-restored_s008-p03-pp02.png"
    $script:HookRejectImagePath = Join-Path $visualBankRoot "VIS-0021_hook-defects-callout-numbered_s008-p03-pp02.png"
    $script:HookAssemblySchemeImagePath = Join-Path $visualBankRoot "схема крюкойо подвески.png"
    $script:S011CategoryImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_long-cargo-reference.png"
    $script:S011LongCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_long-cargo-reference-v2.png"
    $script:S011PieceCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_piece-cargo-reference.png"
    $script:S011StackableCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_stackable-cargo-reference.png"
    $script:S011BulkCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_bulk-cargo-reference.png"
    $script:S011SemiLiquidCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_semi-liquid-cargo-reference.jfif"
    $script:S011LiquidCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_liquid-cargo-reference.png"
    $script:S011GasCargoImagePath = Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\images\S011_gas-cargo-reference.png"

    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = -1
    $presentation = $pp.Presentations.Open($outputAlias, $false, $false, $false)

    $slides = @{}
    $baseS007 = $presentation.Slides.Item(18)

    $slides["S008"] = New-ThemeSlide -Presentation $presentation -Code "S008" -Title "Кран" -Subtitle "Первый опорный слайд темы 2"
    $hero = Add-TextBox -Slide $slides["S008"] -Left 0.82 -Top 1.72 -Width 5.80 -Height 0.56
    Set-TextShape -Shape $hero -Text "Перед строповкой важно понимать сам кран" -FontSize 26 -Bold $true -Color $TEXT
    $heroBody = Add-TextBox -Slide $slides["S008"] -Left 0.82 -Top 2.34 -Width 5.74 -Height 1.02
    Set-Lines -Shape $heroBody -Lines @(
        "Этот слайд открывает тему 2 и задает опорные направления, к которым дальше можно наращивать материал.",
        "Пока фиксируем три главных блока: типы кранов, параметры и конструктивные особенности."
    ) -FontSize 15 -Color $TEXT
    $s008CardP01 = Add-RouteCard -Slide $slides["S008"] -Left 0.82 -Top 4.10 -Width 3.62 -Height 1.68 -Step "1" -Title "Типы кранов" -Body "стреловые`rбашенные`rмостовые`rкозловые`rпортальные" -Accent $BROWN
    $s008CardP02 = Add-RouteCard -Slide $slides["S008"] -Left 4.86 -Top 4.10 -Width 3.62 -Height 1.68 -Step "2" -Title "Основные параметры крана" -Body "Что важно для понимания возможностей крана перед работой с грузом." -Accent $BLUE
    $s008CardP03 = Add-RouteCard -Slide $slides["S008"] -Left 8.90 -Top 4.10 -Width 3.38 -Height 1.68 -Step "3" -Title "Конструктивные особенности" -Body "Какие элементы конструкции влияют на безопасную организацию стропальных работ." -Accent $GREEN

    $slides["S009"] = New-ThemeSlide -Presentation $presentation -Code "S009" -Title "Тема 2.1. Строповка грузов" -Subtitle "Безопасность начинается до подъема"
    $hero = Add-TextBox -Slide $slides["S009"] -Left 0.78 -Top 1.70 -Width 5.85 -Height 0.62
    Set-TextShape -Shape $hero -Text "Сначала оцени груз" -FontSize 28 -Bold $true -Color $TEXT
    $heroBody = Add-TextBox -Slide $slides["S009"] -Left 0.80 -Top 2.38 -Width 5.62 -Height 1.10
    Set-Lines -Shape $heroBody -Lines @(
        "Безопасная строповка начинается не со спешки и не с сигнала крановщику.",
        "Сначала разбираемся, что именно поднимаем и какая схема подойдет этому грузу."
    ) -FontSize 16 -Color $TEXT
    Add-BulletPanel -Slide $slides["S009"] -Left 7.00 -Top 1.66 -Width 5.45 -Height 3.78 -Title "Что внутри блока" -Lines @(
        "Сведения о грузах",
        "Стропы и приспособления",
        "Правила строповки",
        "Проверка схемы",
        "Типовые ошибки"
    ) | Out-Null
    Add-RouteCard -Slide $slides["S009"] -Left 0.82 -Top 4.28 -Width 1.90 -Height 1.20 -Step "1" -Title "Оцени" -Body "Масса, габариты, центр тяжести." -Accent $BROWN | Out-Null
    Add-RouteCard -Slide $slides["S009"] -Left 2.98 -Top 4.28 -Width 1.90 -Height 1.20 -Step "2" -Title "Выбери" -Body "Подходящий строп и способ зацепки." -Accent $BLUE | Out-Null
    Add-RouteCard -Slide $slides["S009"] -Left 5.14 -Top 4.28 -Width 1.90 -Height 1.20 -Step "3" -Title "Проверь" -Body "Схему и условия до подъема." -Accent $GREEN | Out-Null

    $slides["S010"] = New-ThemeSlide -Presentation $presentation -Code "S010" -Title "Маршрут подтемы" -Subtitle "От анализа груза до проверки понимания"
    $routeIntro = Add-TextBox -Slide $slides["S010"] -Left 0.82 -Top 1.66 -Width 11.40 -Height 0.42
    Set-TextShape -Shape $routeIntro -Text "Внутри подтемы идем по понятной логике: от груза и оснастки к правилам, проверке схемы и мини-тесту." -FontSize 16 -Color $TEXT
    Add-RouteCard -Slide $slides["S010"] -Left 0.82 -Top 2.20 -Width 2.30 -Height 1.42 -Step "1" -Title "Грузы" -Body "Что нужно понять до выбора стропа." -Accent $BROWN | Out-Null
    Add-RouteCard -Slide $slides["S010"] -Left 3.28 -Top 2.20 -Width 2.30 -Height 1.42 -Step "2" -Title "Стропы" -Body "Какие средства используются на практике." -Accent $BLUE | Out-Null
    Add-RouteCard -Slide $slides["S010"] -Left 5.74 -Top 2.20 -Width 2.30 -Height 1.42 -Step "3" -Title "Правила" -Body "Точка зацепки и работа по схеме." -Accent $GREEN | Out-Null
    Add-RouteCard -Slide $slides["S010"] -Left 8.20 -Top 2.20 -Width 2.30 -Height 1.42 -Step "4" -Title "Проверка" -Body "Что перепроверить до сигнала." -Accent $OCHRE | Out-Null
    Add-RouteCard -Slide $slides["S010"] -Left 10.66 -Top 2.20 -Width 1.72 -Height 1.42 -Step "5" -Title "Ошибки" -Body "Где чаще всего срывается безопасность." -Accent $RED | Out-Null
    Add-BulletPanel -Slide $slides["S010"] -Left 0.82 -Top 4.02 -Width 11.56 -Height 1.84 -Title "Зачем такая последовательность" -Lines @(
        "Сначала понимаем сам груз.",
        "Потом выбираем средство строповки.",
        "После этого проверяем, можно ли поднимать безопасно."
    ) | Out-Null

    $slides["S011"] = New-ThemeSlide -Presentation $presentation -Code "S011" -Title "Сведения о грузах" -Subtitle "Один груз — не одно и то же решение"
    $s011Lead = Add-TextBox -Slide $slides["S011"] -Left 0.82 -Top 1.70 -Width 11.42 -Height 0.36
    Set-TextShape -Shape $s011Lead -Text "Основные категории грузов, с которыми дальше работаем в подтеме." -FontSize 16 -Color $TEXT
    $s011CardP01 = Add-RouteCard -Slide $slides["S011"] -Left 0.82 -Top 2.20 -Width 2.70 -Height 1.66 -Step "А" -Title "Габаритные" -Body "Обычные грузы с понятными размерами." -Accent $BROWN
    $s011CardP02 = Add-RouteCard -Slide $slides["S011"] -Left 3.74 -Top 2.20 -Width 2.70 -Height 1.66 -Step "Б" -Title "Длинномерные" -Body "Требуют контроля баланса и вылета." -Accent $BLUE
    $s011CardP03 = Add-RouteCard -Slide $slides["S011"] -Left 6.66 -Top 2.20 -Width 2.70 -Height 1.66 -Step "В" -Title "Штучные нештабелируемые" -Body "Отдельные грузы без устойчивой укладки." -Accent $GREEN
    $s011CardP04 = Add-RouteCard -Slide $slides["S011"] -Left 9.58 -Top 2.20 -Width 2.70 -Height 1.66 -Step "Г" -Title "Штучные штабелируемые" -Body "Можно укладывать в ярусы." -Accent $RED
    $s011CardP05 = Add-RouteCard -Slide $slides["S011"] -Left 0.82 -Top 4.10 -Width 2.70 -Height 1.66 -Step "Д" -Title "Насыпные" -Body "Сыпучие материалы." -Accent $OCHRE
    $s011CardP06 = Add-RouteCard -Slide $slides["S011"] -Left 3.74 -Top 4.10 -Width 2.70 -Height 1.66 -Step "Е" -Title "Полужидкие и пластичные" -Body "Густые смеси и пластичные массы." -Accent $STEEL
    $s011CardP07 = Add-RouteCard -Slide $slides["S011"] -Left 6.66 -Top 4.10 -Width 2.70 -Height 1.66 -Step "Ж" -Title "Жидкие" -Body "Перевозятся в емкостях." -Accent $BLUE
    $s011CardP08 = Add-RouteCard -Slide $slides["S011"] -Left 9.58 -Top 4.10 -Width 2.70 -Height 1.66 -Step "З" -Title "Газообразные" -Body "Баллоны и сосуды под давлением." -Accent $GREEN
    $s011HotspotP01 = Add-Hotspot -Slide $slides["S011"] -Left 0.82 -Top 2.20 -Width 2.70 -Height 1.66
    $s011HotspotP02 = Add-Hotspot -Slide $slides["S011"] -Left 3.74 -Top 2.20 -Width 2.70 -Height 1.66
    $s011HotspotP03 = Add-Hotspot -Slide $slides["S011"] -Left 6.66 -Top 2.20 -Width 2.70 -Height 1.66
    $s011HotspotP04 = Add-Hotspot -Slide $slides["S011"] -Left 9.58 -Top 2.20 -Width 2.70 -Height 1.66
    $s011HotspotP05 = Add-Hotspot -Slide $slides["S011"] -Left 0.82 -Top 4.10 -Width 2.70 -Height 1.66
    $s011HotspotP06 = Add-Hotspot -Slide $slides["S011"] -Left 3.74 -Top 4.10 -Width 2.70 -Height 1.66
    $s011HotspotP07 = Add-Hotspot -Slide $slides["S011"] -Left 6.66 -Top 4.10 -Width 2.70 -Height 1.66
    $s011HotspotP08 = Add-Hotspot -Slide $slides["S011"] -Left 9.58 -Top 4.10 -Width 2.70 -Height 1.66

    $slides["S012"] = New-ThemeSlide -Presentation $presentation -Code "S012" -Title "Ключевые факторы груза" -Subtitle "Что влияет на устойчивость и выбор схемы"
    Add-BulletPanel -Slide $slides["S012"] -Left 0.82 -Top 1.66 -Width 5.74 -Height 4.80 -Title "Факторы, которые нельзя пропустить" -Lines @(
        "Масса груза",
        "Габариты",
        "Центр тяжести",
        "Маркировка и знаки",
        "Неизвестная масса - стоп"
    ) | Out-Null
    Add-Panel -Slide $slides["S012"] -Left 6.94 -Top 1.66 -Width 5.48 -Height 4.80 -Title "Задача для ролика" -AccentColor $BLUE -HeaderWidth 2.68 | Out-Null
    $s012Task = Add-TextBox -Slide $slides["S012"] -Left 7.28 -Top 2.28 -Width 4.82 -Height 3.72
    Set-TextShape -Shape $s012Task -Text "Собрать видеоролик, в котором озвучить: Масса груза, габариты, центр тяжести, маркировки и знаки, неизвестная масса." -FontSize 19 -Color $TEXT

    $slides["S013"] = New-ThemeSlide -Presentation $presentation -Code "S013" -Title "Стропы и приспособления" -Subtitle "Средство выбирают под задачу, а не на глаз"
    Add-BulletPanel -Slide $slides["S013"] -Left 0.82 -Top 1.66 -Width 3.00 -Height 4.80 -Title "Что нужно различать" -Lines @(
        "Канатные стропы",
        "Цепные стропы",
        "Текстильные стропы",
        "Грузозахватные приспособления"
    ) | Out-Null
    Add-RouteCard -Slide $slides["S013"] -Left 4.06 -Top 1.78 -Width 2.05 -Height 2.10 -Step "1" -Title "Канатные" -Body "Универсальный рабочий вариант, важен осмотр и состояние ветвей." -Accent $BROWN | Out-Null
    Add-RouteCard -Slide $slides["S013"] -Left 6.28 -Top 1.78 -Width 2.05 -Height 2.10 -Step "2" -Title "Цепные" -Body "Хороши для жестких условий и тяжелых грузов." -Accent $BLUE | Out-Null
    Add-RouteCard -Slide $slides["S013"] -Left 8.50 -Top 1.78 -Width 2.05 -Height 2.10 -Step "3" -Title "Текстильные" -Body "Нужны там, где важна бережная работа с поверхностью." -Accent $GREEN | Out-Null
    Add-RouteCard -Slide $slides["S013"] -Left 10.72 -Top 1.78 -Width 1.58 -Height 2.10 -Step "4" -Title "Оснастка" -Body "Траверсы, захваты, скобы и другие решения под конкретную задачу." -Accent $OCHRE | Out-Null
    Add-BulletPanel -Slide $slides["S013"] -Left 4.06 -Top 4.16 -Width 8.24 -Height 1.58 -Title "Ключевой вывод" -Lines @(
        "Выбор зависит от массы, формы груза и условий работы.",
        "Перед использованием средство не только выбирают, но и обязательно осматривают."
    ) | Out-Null

    $slides["S014"] = New-ThemeSlide -Presentation $presentation -Code "S014" -Title "Что проверять перед использованием" -Subtitle "Осмотр - это часть допуска к работе"
    Add-BulletPanel -Slide $slides["S014"] -Left 0.82 -Top 1.66 -Width 5.70 -Height 4.80 -Title "Перед подъемом обязательно" -Lines @(
        "Осмотреть строп или приспособление.",
        "Проверить бирку и маркировку.",
        "Убедиться, что нет повреждений и износа.",
        "Неисправное - не использовать."
    ) | Out-Null
    $s013Image = Join-Path (Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\diagrams\error-analysis") $imageMap["S013"]
    Add-ImagePanel -Slide $slides["S014"] -Left 6.94 -Top 1.66 -Width 5.48 -Height 4.80 -Title "Опорный экран осмотра" `
        -ImagePath $s013Image `
        -Caption "Экран осмотра уже можно использовать как рабочий визуальный каркас для линейного объяснения."

    $slides["S015"] = New-ThemeSlide -Presentation $presentation -Code "S015" -Title "Правила строповки" -Subtitle "Правильная зацепка начинается с дисциплины схемы"
    Add-RouteCard -Slide $slides["S015"] -Left 0.82 -Top 1.80 -Width 3.74 -Height 2.06 -Step "1" -Title "Точка зацепки" -Body "Выбираем понятную и безопасную точку, а не случайный удобный хват." -Accent $BROWN | Out-Null
    Add-RouteCard -Slide $slides["S015"] -Left 4.78 -Top 1.80 -Width 3.74 -Height 2.06 -Step "2" -Title "Схема строповки" -Body "Работаем по схеме, а не по привычке или похожему варианту." -Accent $GREEN | Out-Null
    Add-RouteCard -Slide $slides["S015"] -Left 8.74 -Top 1.80 -Width 3.54 -Height 2.06 -Step "3" -Title "Без случайной обвязки" -Body "Если схема вызывает сомнение, ее нельзя считать допустимой." -Accent $RED | Out-Null
    Add-BulletPanel -Slide $slides["S015"] -Left 0.82 -Top 4.22 -Width 11.46 -Height 1.72 -Title "Главный смысл экрана" -Lines @(
        "Безопасность рушится в тот момент, когда зацепку принимают на глаз.",
        "Правило простое: понятная точка, проверенная схема, никакой самодеятельности."
    ) | Out-Null

    $slides["S016"] = New-ThemeSlide -Presentation $presentation -Code "S016" -Title "Запреты и риски" -Subtitle "Где схема внешне выглядит рабочей, но уже опасна"
    $s015Image = Join-Path (Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\diagrams\error-analysis") $imageMap["S015"]
    Add-ImagePanel -Slide $slides["S016"] -Left 0.82 -Top 1.66 -Width 5.66 -Height 4.80 -Title "Нагрузка на ветви" `
        -ImagePath $s015Image `
        -Caption "Готовая схема помогает объяснить, почему угол между ветвями нельзя игнорировать."
    Add-BulletPanel -Slide $slides["S016"] -Left 6.92 -Top 1.66 -Width 5.50 -Height 4.80 -Title "Что нельзя пропускать" -Lines @(
        "Угол между ветвями влияет на нагрузку.",
        "Острые ребра требуют защиты стропа.",
        "Опасная зацепка не становится нормой, даже если раньше так делали.",
        "При риске - стоп."
    ) | Out-Null

    $slides["S017"] = New-ThemeSlide -Presentation $presentation -Code "S017" -Title "Проверка схемы строповки" -Subtitle "Перед сигналом должна быть короткая, но обязательная проверка"
    Add-Panel -Slide $slides["S017"] -Left 0.82 -Top 1.66 -Width 5.74 -Height 4.80 -Title "Проверка по шагам" | Out-Null
    Add-StepRow -Slide $slides["S017"] -Left 1.08 -Top 2.26 -Width 5.18 -IndexText "1" -Text "Проверь массу груза" | Out-Null
    Add-StepRow -Slide $slides["S017"] -Left 1.08 -Top 3.04 -Width 5.18 -IndexText "2" -Text "Проверь схему строповки" | Out-Null
    Add-StepRow -Slide $slides["S017"] -Left 1.08 -Top 3.82 -Width 5.18 -IndexText "3" -Text "Проверь устойчивость" | Out-Null
    Add-StepRow -Slide $slides["S017"] -Left 1.08 -Top 4.60 -Width 5.18 -IndexText "4" -Text "Проверь стропы и приспособления" | Out-Null
    $s016Image = Join-Path (Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\diagrams\error-analysis") $imageMap["S016"]
    Add-ImagePanel -Slide $slides["S017"] -Left 6.94 -Top 1.66 -Width 5.48 -Height 4.80 -Title "Опорная последовательность" `
        -ImagePath $s016Image `
        -Caption "Сначала схема и проверка, потом сигнал и только затем подъем."

    $slides["S018"] = New-ThemeSlide -Presentation $presentation -Code "S018" -Title "Когда работу нужно остановить" -Subtitle "Стоп - это обязанность, а не пауза из осторожности"
    Add-BulletPanel -Slide $slides["S018"] -Left 0.82 -Top 1.66 -Width 5.72 -Height 4.80 -Title "Сигналы на остановку" -Lines @(
        "Неясна масса груза.",
        "Нет надежной схемы.",
        "Есть риск перекоса или разворота.",
        "Оснастка неисправна.",
        "Опасность заметили раньше, чем груз пошел в движение."
    ) | Out-Null
    $s017Image = Join-Path (Join-Path $aliasRoot "assets\course-media\module-01-stropovka-gruzov\diagrams\error-analysis") $imageMap["S017"]
    Add-ImagePanel -Slide $slides["S018"] -Left 6.96 -Top 1.66 -Width 5.46 -Height 4.80 -Title "Опорный visual asset" `
        -ImagePath $s017Image `
        -Caption "Подъем не продолжают на сомнительной схеме: сначала устраняют риск, потом возобновляют работу."

    $slides["S019"] = New-ThemeSlide -Presentation $presentation -Code "S019" -Title "Типовые схемы и ошибки" -Subtitle "Переход от правил к распознаванию опасных вариантов"
    Add-RouteCard -Slide $slides["S019"] -Left 0.82 -Top 1.78 -Width 2.74 -Height 1.64 -Step "1" -Title "Длинномерные" -Body "Важно удержать баланс и исключить раскачивание." -Accent $BROWN | Out-Null
    Add-RouteCard -Slide $slides["S019"] -Left 3.78 -Top 1.78 -Width 2.74 -Height 1.64 -Step "2" -Title "Железобетон" -Body "Смотрим на массу, опоры и надежность зацепки." -Accent $BLUE | Out-Null
    Add-RouteCard -Slide $slides["S019"] -Left 6.74 -Top 1.78 -Width 2.74 -Height 1.64 -Step "3" -Title "Листовой металл" -Body "Особенно важны кромки и защита стропа." -Accent $GREEN | Out-Null
    Add-RouteCard -Slide $slides["S019"] -Left 9.70 -Top 1.78 -Width 2.54 -Height 1.64 -Step "4" -Title "Трубы и оборудование" -Body "Нужна проверенная схема и контроль устойчивости." -Accent $OCHRE | Out-Null
    Add-BulletPanel -Slide $slides["S019"] -Left 0.82 -Top 4.04 -Width 11.42 -Height 1.88 -Title "Что должен сделать этот экран" -Lines @(
        "Связать правила темы с реальными примерами схем.",
        "Подвести слушателя к проверке понимания: где ошибка и почему она опасна.",
        "Сделать мост к тесту S020."
    ) | Out-Null

    $slides["S020"] = New-ThemeSlide -Presentation $presentation -Code "S020" -Title "Промежуточный тест" -Subtitle "Проверка по теме 2.1"
    Add-BulletPanel -Slide $slides["S020"] -Left 0.82 -Top 1.66 -Width 5.76 -Height 4.80 -Title "Что проверяем" -Lines @(
        "Какой груз перед нами и что важно о нем знать.",
        "Какой строп или приспособление подходит под задачу.",
        "Какая схема опасна и когда работу нужно остановить.",
        "Насколько уверенно слушатель распознает типовые ошибки."
    ) | Out-Null
    Add-Panel -Slide $slides["S020"] -Left 6.96 -Top 1.66 -Width 5.42 -Height 4.80 -Title "Экран входа в тест" | Out-Null
    $qTitle = Add-TextBox -Slide $slides["S020"] -Left 7.30 -Top 2.24 -Width 4.72 -Height 0.42
    Set-TextShape -Shape $qTitle -Text "Категории вопросов" -FontSize 20 -Bold $true -Color $TEXT
    $qList = Add-TextBox -Slide $slides["S020"] -Left 7.30 -Top 2.82 -Width 4.66 -Height 2.30
    Set-Lines -Shape $qList -Lines @(
        "• проверка груза",
        "• неизвестная масса",
        "• центр тяжести",
        "• угол между ветвями",
        "• оснастка, схема и решение стоп"
    ) -FontSize 18 -Color $TEXT
    $qNote = Add-TextBox -Slide $slides["S020"] -Left 7.30 -Top 5.34 -Width 4.72 -Height 0.64
    Set-Lines -Shape $qNote -Lines @(
        "В этой итерации собираем линейный вход в тест.",
        "Разборы ошибок добавим следующим слоем по мере сборки."
    ) -FontSize 12 -Color $MUTED

    $slides["S021"] = New-ThemeSlide -Presentation $presentation -Code "S021" -Title "Финальный вывод по подтеме" -Subtitle "Короткая рабочая последовательность перед подъемом"
    Add-RouteCard -Slide $slides["S021"] -Left 0.82 -Top 1.92 -Width 2.74 -Height 1.70 -Step "1" -Title "Оцени груз" -Body "Масса, габариты, центр тяжести и маркировка." -Accent $BROWN | Out-Null
    Add-RouteCard -Slide $slides["S021"] -Left 3.78 -Top 1.92 -Width 2.74 -Height 1.70 -Step "2" -Title "Выбери строп" -Body "Средство и схему под конкретную задачу." -Accent $BLUE | Out-Null
    Add-RouteCard -Slide $slides["S021"] -Left 6.74 -Top 1.92 -Width 2.74 -Height 1.70 -Step "3" -Title "Проверь схему" -Body "До сигнала, а не после пробного подъема." -Accent $GREEN | Out-Null
    Add-RouteCard -Slide $slides["S021"] -Left 9.70 -Top 1.92 -Width 2.54 -Height 1.70 -Step "4" -Title "Не рискуй" -Body "При сомнении останавливай работу." -Accent $RED | Out-Null
    Add-BulletPanel -Slide $slides["S021"] -Left 0.82 -Top 4.18 -Width 11.42 -Height 1.74 -Title "Итоговый тезис" -Lines @(
        "Безопасная строповка - это не отдельный прием, а последовательность решений до подъема.",
        "Сомнение в схеме или данных о грузе - достаточная причина не продолжать работу."
    ) | Out-Null

    $s007Next = Add-Button -Slide $baseS007 -Left 9.34 -Top 6.10 -Width 3.02 -Height 0.50 -Text "К теме 2.1" -FillColor $BROWN -FontSize 16
    Set-SlideJump -Shape $s007Next -TargetSlide $slides["S008"]

    $slides["S008-P01-PP01"] = New-CranePurposeSlide -Presentation $presentation -Code "S008-P01-PP01" -BackTarget $null
    $slides["S008-P01"] = New-CraneTypesSlide -Presentation $presentation -Code "S008-P01" -BackTarget $slides["S008"] -DetailTarget $slides["S008-P01-PP01"] -ImageMap @{
        boom = (Join-Path $craneImageRoot "mobile-boom-crane.png")
        tower = (Join-Path $craneImageRoot "tower-crane.png")
        bridge = (Join-Path $craneImageRoot "bridge-crane.png")
        gantry = (Join-Path $craneImageRoot "gantry-crane.png")
        manipulator = (Join-Path $craneImageRoot "loader-crane-manipulator.png")
        pipe = (Join-Path $craneImageRoot "pipe-layer-crane.png")
    }
    Set-SlideJump -Shape $slides["S008-P01-PP01"].Shapes.Item($slides["S008-P01-PP01"].Shapes.Count) -TargetSlide $slides["S008-P01"]
    $slides["S008-P02"] = New-CraneParametersSlide -Presentation $presentation -Code "S008-P02" -BackTarget $slides["S008"]
    $slides["S008-P03-PP01"] = New-HookAssemblySlide -Presentation $presentation -Code "S008-P03-PP01" -BackTarget $null
    $slides["S008-P03-PP02"] = New-HookDetailSlide -Presentation $presentation -Code "S008-P03-PP02" -BackTarget $null
    $slides["S008-P03"] = New-CraneConstructionSlide -Presentation $presentation -Code "S008-P03" -BackTarget $slides["S008"] -HookTarget $slides["S008-P03-PP02"] -HookAssemblyTarget $slides["S008-P03-PP01"]
    $slides["S011-P01"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P01" -Title "Габаритные грузы" -ItemTitle "Габаритный груз" -Definition "это груз, размеры которого не превышают допустимые нормы для перевозки на стандартном транспорте и не создают помех при движении по дорогам общего пользования. Для такого груза не нужны специальные разрешения или сопровождение." -ImagePath $script:S011CategoryImagePath -BackTarget $slides["S011"] -AccentColor $BROWN
    $slides["S011-P02"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P02" -Title "Длинномерные грузы" -ItemTitle "Длинномерный груз" -Definition "это груз, длина которого заметно превышает его ширину и высоту. При подъеме такой груз особенно чувствителен к перекосу, раскачиванию и смещению центра тяжести, поэтому требует устойчивой схемы строповки, правильной расстановки точек зацепки и постоянного контроля баланса." -ImagePath $script:S011LongCargoImagePath -BackTarget $slides["S011"] -AccentColor $BLUE
    $slides["S011-P03"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P03" -Title "Штучные нештабелируемые грузы" -ItemTitle "Штучный нештабелируемый груз" -Definition "это отдельный груз, который нельзя безопасно укладывать в ярусы или устойчиво складировать без специальной оснастки. Для него особенно важно исключить перекос, случайное смещение и потерю устойчивости во время подъема и перемещения." -ImagePath $script:S011PieceCargoImagePath -BackTarget $slides["S011"] -AccentColor $GREEN -BodyText "• Штучные нештабелируемые грузы — это отдельные крупногабаритные, тяжелые или нестандартные по форме предметы (оборудование, станки, металлоконструкции), которые из-за сложной геометрии или хрупкости нельзя ставить друг на друга в несколько ярусов."
    $slides["S011-P04"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P04" -Title "Штучные штабелируемые грузы" -ItemTitle "Штучный штабелируемый груз" -Definition "это отдельные грузы, которые допускается устойчиво укладывать в несколько ярусов при хранении и перемещении. При строповке для них важно сохранить устойчивость штабеля, не допустить смещения верхних рядов и учитывать, как распределяется масса по всей пачке или укладке." -ImagePath $script:S011StackableCargoImagePath -BackTarget $slides["S011"] -AccentColor $RED -BodyText "Штучные штабелируемые грузы — это одинаковые по форме предметы (ящики, контейнеры, поддоны с кирпичом, трубы, плиты), которые имеют ровные грани или специальные пазы.`r`n`r`nИх главная особенность — их можно безопасно складывать друг на друга в несколько ярусов (в штабели) для компактного хранения." -BodyFontSize 16
    $slides["S011-P05"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P05" -Title "Насыпные грузы" -ItemTitle "Насыпной груз" -Definition "это груз из множества мелких однородных частиц, который не сохраняет собственной формы и свободно пересыпается. При его подъеме и перемещении важно учитывать текучесть, возможность смещения массы внутри тары и риск просыпания." -ImagePath $script:S011BulkCargoImagePath -BackTarget $slides["S011"] -AccentColor $OCHRE -BodyText "• Насыпные грузы — это однородные сыпучие материалы (песок, щебень, уголь, зерно), которые перевозятся и хранятся без упаковки, навалом. Их перемещают не стропами, а специальными грузозахватными устройствами — грейферами и тарой." -BodyFontSize 16
    $slides["S011-P06"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P06" -Title "Полужидкие и пластичные грузы" -ItemTitle "Полужидкий или пластичный груз" -Definition "это вязкие, густые или легко деформируемые материалы, которые частично сохраняют форму, но под нагрузкой могут растекаться, сминаться или смещаться внутри тары. При перемещении для них особенно важно учитывать устойчивость упаковки, герметичность емкости и поведение массы при качке." -ImagePath $script:S011SemiLiquidCargoImagePath -BackTarget $slides["S011"] -AccentColor $STEEL -BodyText "Полужидкие и пластичные грузы — это вязкие массы (бетон, строительный раствор, битум, мастика), которые из-за своей текучести не имеют постоянной формы. Для их перемещения краном стропальщики используют специальную тару — бадьи или бункеры." -BodyFontSize 16
    $slides["S011-P07"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P07" -Title "Жидкие грузы" -ItemTitle "Жидкий груз" -Definition "это вещества, которые полностью принимают форму емкости и при перемещении могут переливаться, создавать гидроудары и смещать центр тяжести. Для безопасного подъема важно использовать герметичную тару и учитывать колебание жидкости внутри сосуда." -ImagePath $script:S011LiquidCargoImagePath -BackTarget $slides["S011"] -AccentColor $BLUE -BodyText "Жидкие грузы — это текучие жидкости (вода, топливо, кислоты, щелочи), которые перевозятся только в герметичной таре — бочках, цистернах или бутылях. Их главная опасность — при движении они колеблются, смещают центр тяжести и раскачивают кран, поэтому тару никогда не заливают до самого верха." -BodyFontSize 16
    $slides["S011-P08"] = New-CargoCategoryDetailSlide -Presentation $presentation -Code "S011-P08" -Title "Газообразные грузы" -ItemTitle "Газообразный груз" -Definition "это газы, которые перевозятся и хранятся только в герметичных баллонах, сосудах или специальных емкостях под давлением. При их перемещении особенно важно исключить удары, перегрев, падение тары и любое нарушение герметичности." -ImagePath $script:S011GasCargoImagePath -BackTarget $slides["S011"] -AccentColor $GREEN -BodyText "Газообразные грузы — это газы (кислород, пропан, ацетилен, азот), которые сжаты под огромным давлением и перемещаются исключительно в специальных прочных баллонах или контейнерах." -BodyFontSize 16
    Set-SlideJump -Shape $slides["S008-P03-PP01"].Shapes.Item($slides["S008-P03-PP01"].Shapes.Count) -TargetSlide $slides["S008-P03"]
    Set-SlideJump -Shape $slides["S008-P03-PP02"].Shapes.Item($slides["S008-P03-PP02"].Shapes.Count) -TargetSlide $slides["S008-P03"]
    $slides["S008-P01"].MoveTo(20)
    $slides["S008-P01-PP01"].MoveTo(21)
    $slides["S008-P02"].MoveTo(22)
    $slides["S008-P03"].MoveTo(23)
    $slides["S008-P03-PP01"].MoveTo(24)
    $slides["S008-P03-PP02"].MoveTo(25)
    $slides["S011-P01"].MoveTo(29)
    $slides["S011-P02"].MoveTo(30)
    $slides["S011-P03"].MoveTo(31)
    $slides["S011-P04"].MoveTo(32)
    $slides["S011-P05"].MoveTo(33)
    $slides["S011-P06"].MoveTo(34)
    $slides["S011-P07"].MoveTo(35)
    $slides["S011-P08"].MoveTo(36)

    $linearCodes = @("S008", "S009", "S010", "S011", "S012", "S013", "S014", "S015", "S016", "S017", "S018", "S019", "S020", "S021")
    for ($i = 0; $i -lt $linearCodes.Count; $i++) {
        $code = $linearCodes[$i]
        $slide = $slides[$code]
        $prevTarget = $null
        $nextTarget = $null
        $prevText = $null
        $nextText = $null

        if ($code -eq "S008") {
            $prevTarget = $baseS007
            $prevText = "Назад"
        }
        elseif ($i -gt 0) {
            $prevTarget = $slides[$linearCodes[$i - 1]]
            $prevText = "Назад"
        }

        if ($i -lt ($linearCodes.Count - 1)) {
            $nextTarget = $slides[$linearCodes[$i + 1]]
            $nextText = "Далее"
        }

        $nav = Add-Nav -Slide $slide -PrevText $prevText -NextText $nextText
        if ($nav[0] -and $prevTarget) {
            Set-SlideJump -Shape $nav[0] -TargetSlide $prevTarget
        }
        if ($nav[1] -and $nextTarget) {
            Set-SlideJump -Shape $nav[1] -TargetSlide $nextTarget
        }
    }

    Set-SlideJump -Shape $s008CardP01 -TargetSlide $slides["S008-P01"]
    Set-SlideJump -Shape $s008CardP02 -TargetSlide $slides["S008-P02"]
    Set-SlideJump -Shape $s008CardP03 -TargetSlide $slides["S008-P03"]
    Set-SlideJump -Shape $s011HotspotP01 -TargetSlide $slides["S011-P01"]
    Set-SlideJump -Shape $s011HotspotP02 -TargetSlide $slides["S011-P02"]
    Set-SlideJump -Shape $s011HotspotP03 -TargetSlide $slides["S011-P03"]
    Set-SlideJump -Shape $s011HotspotP04 -TargetSlide $slides["S011-P04"]
    Set-SlideJump -Shape $s011HotspotP05 -TargetSlide $slides["S011-P05"]
    Set-SlideJump -Shape $s011HotspotP06 -TargetSlide $slides["S011-P06"]
    Set-SlideJump -Shape $s011HotspotP07 -TargetSlide $slides["S011-P07"]
    Set-SlideJump -Shape $s011HotspotP08 -TargetSlide $slides["S011-P08"]
    $presentation.Save()
}
finally {
    if ($presentation) {
        $presentation.Close()
    }
    if ($pp) {
        $pp.Quit()
    }
    if ($substDrive) {
        & subst $substDrive /d | Out-Null
    }
}

Write-Output $outputPath
