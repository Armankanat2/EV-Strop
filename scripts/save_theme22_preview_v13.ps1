$ErrorActionPreference = "Stop"

$path = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S029_theme22_preview_2026-06-30_v13_ev-group-techstyle-premium-subslides.pptx"

Add-Type -AssemblyName Microsoft.VisualBasic
$ppt = [Microsoft.VisualBasic.Interaction]::GetObject("", "PowerPoint.Application")
$saved = $false

foreach ($presentation in $ppt.Presentations) {
    if ($presentation.FullName -eq $path) {
        $presentation.Save()
        $saved = $true
        break
    }
}

if (-not $saved) {
    $presentation = $ppt.Presentations.Open($path, $false, $false, $false)
    $presentation.Save()
    $presentation.Close()
}

Write-Output "saved"
