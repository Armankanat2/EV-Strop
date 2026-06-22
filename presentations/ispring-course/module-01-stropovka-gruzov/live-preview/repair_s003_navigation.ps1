$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

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

function Repair-S003Deck {
    param(
        [Parameter(Mandatory = $true)] [string] $PptxName,
        [Parameter(Mandatory = $true)] [int] $MainSlideIndex,
        [Parameter(Mandatory = $true)] [int[]] $DetailSlideIndexes,
        [Parameter(Mandatory = $true)] [int] $BackTargetIndex,
        [Parameter(Mandatory = $true)] [int] $NextTargetIndex
    )

    $pptxPath = Join-Path $scriptDir $PptxName
    $ppsxPath = [System.IO.Path]::ChangeExtension($pptxPath, ".ppsx")

    if (-not (Test-Path $pptxPath)) {
        throw "Presentation not found: $pptxPath"
    }

    $pres = $pp.Presentations.Open($pptxPath, $false, $false, $false)
    $main = $pres.Slides.Item($MainSlideIndex)
    $targets = @(
        @{ ShapeIndex = 9; Target = $pres.Slides.Item($DetailSlideIndexes[0]) },
        @{ ShapeIndex = 12; Target = $pres.Slides.Item($DetailSlideIndexes[1]) },
        @{ ShapeIndex = 15; Target = $pres.Slides.Item($DetailSlideIndexes[2]) },
        @{ ShapeIndex = 18; Target = $pres.Slides.Item($DetailSlideIndexes[3]) }
    )

    foreach ($item in $targets) {
        Set-SlideJump -Shape $main.Shapes.Item($item.ShapeIndex) -TargetSlide $item.Target
    }

    Set-SlideJump -Shape $main.Shapes.Item(24) -TargetSlide $pres.Slides.Item($BackTargetIndex)
    Set-SlideJump -Shape $main.Shapes.Item(25) -TargetSlide $pres.Slides.Item($NextTargetIndex)

    foreach ($detailIndex in $DetailSlideIndexes) {
        $detail = $pres.Slides.Item($detailIndex)
        Set-SlideJump -Shape $detail.Shapes.Item(14) -TargetSlide $main
    }

    $pres.Save()
    if (Test-Path $ppsxPath) {
        Remove-Item -LiteralPath $ppsxPath -Force
    }
    $pres.SaveAs($ppsxPath)
    $pres.Close()

    Write-Output $pptxPath
    Write-Output $ppsxPath
}

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1

Repair-S003Deck `
    -PptxName "S003_npa_preview_2026-06-22_v4.pptx" `
    -MainSlideIndex 1 `
    -DetailSlideIndexes @(2, 3, 4, 5) `
    -BackTargetIndex 1 `
    -NextTargetIndex 2

Repair-S003Deck `
    -PptxName "S001-S007_live_preview_working_2026-06-22_v4.pptx" `
    -MainSlideIndex 5 `
    -DetailSlideIndexes @(6, 7, 8, 9) `
    -BackTargetIndex 2 `
    -NextTargetIndex 10

$pp.Quit()
