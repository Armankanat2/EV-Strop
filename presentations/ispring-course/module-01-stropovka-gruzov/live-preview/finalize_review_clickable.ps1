param(
    [string] $SourcePptxPath,
    [string] $ReviewStem,
    [switch] $Open
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir "..\..\..\..\scripts\powerpoint_stable_tools.ps1")

if (-not $SourcePptxPath) {
    $latestDeck = Get-ChildItem -LiteralPath $scriptDir -Filter "S001-S007_live_preview_working_*.pptx" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $latestDeck) {
        throw "Could not find a live preview PPTX to finalize in $scriptDir"
    }

    $sourceItem = $latestDeck
}
else {
    $sourceItem = Get-Item -LiteralPath $SourcePptxPath
}

if (-not $ReviewStem) {
    if ($sourceItem.BaseName -match '^S001-S007_live_preview_working_(\d{4}-\d{2}-\d{2})_v\d+$') {
        $ReviewStem = "S001-S007_review_clickable_$($Matches[1])_v01"
    }
    else {
        $ReviewStem = $sourceItem.BaseName -replace "live_preview_working", "review_clickable"
    }
}

$reviewPptxPath = Join-Path $sourceItem.DirectoryName ($ReviewStem + ".pptx")
$reviewPpsxPath = Join-Path $sourceItem.DirectoryName ($ReviewStem + ".ppsx")

Copy-Item -LiteralPath $sourceItem.FullName -Destination $reviewPptxPath -Force
Export-StablePpsx -SourcePath $reviewPptxPath -OutputPath $reviewPpsxPath | Out-Null

if ($Open) {
    Open-PowerPointFile -Path $reviewPpsxPath | Out-Null
}

Write-Output $reviewPptxPath
Write-Output $reviewPpsxPath
