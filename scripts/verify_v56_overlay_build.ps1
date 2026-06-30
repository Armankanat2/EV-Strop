$ErrorActionPreference = "Stop"

$src = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-29_v55.pptx"
$dst = Join-Path $PWD "presentations/ispring-course/module-01-stropovka-gruzov/live-preview/S001-S040_live_preview_working_2026-06-30_v56_ev-group-overlay.pptx"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Count-Media {
    param([string]$path)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        return ($zip.Entries | Where-Object { $_.FullName -like "ppt/media/*" }).Count
    }
    finally {
        $zip.Dispose()
    }
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$presentation = $ppt.Presentations.Open($dst, $true, $false, $false)

try {
    Write-Output ("src_media={0}" -f (Count-Media $src))
    Write-Output ("dst_media={0}" -f (Count-Media $dst))
    Write-Output ("slides={0}" -f $presentation.Slides.Count)
}
finally {
    $presentation.Close()
    $ppt.Quit()
}
