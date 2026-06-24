param(
    [Parameter(Mandatory = $true)]
    [string] $FilePath
)

$ErrorActionPreference = "Stop"

. (Join-Path $PSScriptRoot "powerpoint_stable_tools.ps1")

$resolvedPath = (Resolve-Path -LiteralPath $FilePath).Path
$extension = [System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant()
$pathToOpen = $resolvedPath

if ($extension -eq ".ppsx") {
    $kind = Get-DeckPackageKind -Path $resolvedPath

    if ($kind -ne "slideshow") {
        $siblingPptx = [System.IO.Path]::ChangeExtension($resolvedPath, ".pptx")
        $conversionSource = $resolvedPath

        if (Test-Path -LiteralPath $siblingPptx) {
            $conversionSource = $siblingPptx
        }

        $pathToOpen = Convert-ToStablePpsx -SourcePath $conversionSource
    }
}

Open-PowerPointFile -Path $pathToOpen | Out-Null
Write-Output $pathToOpen
