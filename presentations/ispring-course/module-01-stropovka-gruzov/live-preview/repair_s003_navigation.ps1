param(
    [ValidateSet("all", "s003", "combined")]
    [string] $TargetDeck = "all"
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "..\..\..\..\scripts\powerpoint_stable_tools.ps1")

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

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

function Get-ShapeByName {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [string] $Name
    )

    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        if ($shape.Name -eq $Name) {
            return $shape
        }
    }

    throw "Shape with name '$Name' not found on slide $($Slide.SlideIndex)"
}

function Get-OrCreateInfoGroup {
    param(
        [Parameter(Mandatory = $true)] $Slide
    )

    for ($i = 1; $i -le $Slide.Shapes.Count; $i++) {
        $shape = $Slide.Shapes.Item($i)
        if ($shape.Name -eq "S003_INFO_GROUP") {
            return $shape
        }
    }

    $group = $Slide.Shapes.Range(@("S003_INFO_PANEL", "S003_INFO_HEAD", "S003_INFO_BODY")).Group()
    $group.Name = "S003_INFO_GROUP"
    return $group
}

function Open-PresentationSafe {
    param(
        [Parameter(Mandatory = $true)] [string] $PptxPath
    )

    try {
        return $pp.Presentations.Open($PptxPath, $false, $false, $false)
    }
    catch {
        throw "PowerPoint automation could not open '$PptxPath'. The file may be open in another PowerPoint window or need a fresh manual Save As. Original error: $($_.Exception.Message)"
    }
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

    $pres = Open-PresentationSafe -PptxPath $pptxPath
    $main = $pres.Slides.Item($MainSlideIndex)
    $targets = @(
        @{ Shape = (Get-ShapeByName -Slide $main -Name "S003_LAW_LINK_01"); Target = $pres.Slides.Item($DetailSlideIndexes[0]) },
        @{ Shape = (Get-ShapeByName -Slide $main -Name "S003_LAW_LINK_02"); Target = $pres.Slides.Item($DetailSlideIndexes[1]) },
        @{ Shape = (Get-ShapeByName -Slide $main -Name "S003_LAW_LINK_03"); Target = $pres.Slides.Item($DetailSlideIndexes[2]) },
        @{ Shape = (Get-ShapeByName -Slide $main -Name "S003_LAW_LINK_04"); Target = $pres.Slides.Item($DetailSlideIndexes[3]) }
    )

    foreach ($item in $targets) {
        Set-SlideJump -Shape $item.Shape -TargetSlide $item.Target
    }

    Set-SlideJump -Shape (Get-ShapeByName -Slide $main -Name "S003_BACK") -TargetSlide $pres.Slides.Item($BackTargetIndex)
    Set-SlideJump -Shape (Get-ShapeByName -Slide $main -Name "S003_NEXT") -TargetSlide $pres.Slides.Item($NextTargetIndex)

    $detailBackNames = @("S003_P01_BACK", "S003_P02_BACK", "S003_P03_BACK", "S003_P04_BACK")
    for ($i = 0; $i -lt $DetailSlideIndexes.Count; $i++) {
        $detailIndex = $DetailSlideIndexes[$i]
        $detail = $pres.Slides.Item($detailIndex)
        Set-SlideJump -Shape (Get-ShapeByName -Slide $detail -Name $detailBackNames[$i]) -TargetSlide $main
    }

    $infoGroup = Get-OrCreateInfoGroup -Slide $main
    $sequence = $main.TimeLine.MainSequence

    for ($i = $sequence.Count; $i -ge 1; $i--) {
        $effect = $sequence.Item($i)
        if ($effect.Shape.Id -eq $infoGroup.Id) {
            $effect.Delete()
        }
    }

    $panelEffect = $sequence.AddEffect(
        $infoGroup,
        [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade,
        0,
        [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerAfterPrevious
    )
    $panelEffect.Timing.TriggerDelayTime = 2
    $panelEffect.Timing.Duration = 0.9

    $panelExit = $sequence.AddEffect(
        $infoGroup,
        [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade,
        0,
        [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerAfterPrevious
    )
    $panelExit.Exit = $true
    $panelExit.Timing.TriggerDelayTime = 5
    $panelExit.Timing.Duration = 0.9

    $pres.Save()
    $pres.Close()
    Export-StablePpsx -SourcePath $pptxPath -OutputPath $ppsxPath | Out-Null

    Write-Output $pptxPath
    Write-Output $ppsxPath
}

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1

$deckConfigs = @(
    @{
        Key = "s003"
        PptxName = "S003_npa_preview_2026-06-22_v4.pptx"
        MainSlideIndex = 1
        DetailSlideIndexes = @(2, 3, 4, 5)
        BackTargetIndex = 1
        NextTargetIndex = 2
    },
    @{
        Key = "combined"
        PptxName = "S001-S007_live_preview_working_2026-06-22_v4.pptx"
        MainSlideIndex = 5
        DetailSlideIndexes = @(6, 7, 8, 9)
        BackTargetIndex = 2
        NextTargetIndex = 10
    }
)

if ($TargetDeck -ne "all") {
    $deckConfigs = $deckConfigs | Where-Object { $_.Key -eq $TargetDeck }
}

$failures = @()

try {
    foreach ($deck in $deckConfigs) {
        try {
            Repair-S003Deck `
                -PptxName $deck.PptxName `
                -MainSlideIndex $deck.MainSlideIndex `
                -DetailSlideIndexes $deck.DetailSlideIndexes `
                -BackTargetIndex $deck.BackTargetIndex `
                -NextTargetIndex $deck.NextTargetIndex
        }
        catch {
            $message = "$($deck.Key): $($_.Exception.Message)"
            Write-Warning $message
            $failures += $message
        }
    }
}
finally {
    if ($pp) {
        $pp.Quit()
    }
}

if ($failures.Count -gt 0) {
    throw "Repair completed with failures:`n$($failures -join "`n")"
}
