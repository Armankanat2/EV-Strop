$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$baseBuilder = Join-Path $scriptDir "build_s001_s002_live_preview.py"
$pptxPath = Join-Path $scriptDir "S001-S002_live_preview_working.pptx"
$ppsxPath = Join-Path $scriptDir "S001-S002_live_preview_working.ppsx"
$mp4Path = Join-Path $scriptDir "video-build\s002-p01-preview-2026-06-20\S002-P01_video_preview_61_14sec_2026-06-20.mp4"
$wmvPath = Join-Path $scriptDir "video-build\s002-p01-preview-2026-06-20\S002-P01_video_preview_61_14sec_2026-06-20.wmv"
$ffmpegPath = "C:\Program Files (x86)\ClipGrab\ffmpeg.exe"

if (-not (Test-Path $mp4Path)) {
    throw "Video file not found: $mp4Path"
}

if (-not (Test-Path $ffmpegPath)) {
    throw "FFmpeg not found: $ffmpegPath"
}

python $baseBuilder | Out-Null

& $ffmpegPath -y -i $mp4Path -c:v wmv2 -b:v 1800k -c:a wmav2 -b:a 128k $wmvPath | Out-Null

$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = -1

try {
    $pres = $pp.Presentations.Open($pptxPath, $false, $false, $false)
    $slide = $pres.Slides.Item(3)

    foreach ($shapeIndex in @(7, 6)) {
        if ($slide.Shapes.Count -ge $shapeIndex) {
            $shape = $slide.Shapes.Item($shapeIndex)
            if ($shape.Type -eq 17) {
                $shape.Delete()
            }
        }
    }

    for ($i = $slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $slide.Shapes.Item($i)

        if ($shape.Type -eq 16) {
            $shape.Delete()
            continue
        }
    }

    $left = 0.82 * 72
    $top = 1.89 * 72
    $width = 4.56 * 72
    $height = 2.565 * 72

    $media = $slide.Shapes.AddMediaObject($wmvPath, $left, $top, $width, $height)
    $media.Name = "S002_P01_VIDEO"

    $playSettings = $media.AnimationSettings.PlaySettings
    $playSettings.PlayOnEntry = -1
    $playSettings.HideWhileNotPlaying = 0
    $playSettings.LoopUntilStopped = 0
    $playSettings.PauseAnimation = 0
    $playSettings.RewindMovie = -1
    $playSettings.StopAfterSlides = 1

    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if (
            $shape.Type -eq 1 -and
            [math]::Abs($shape.Left - (1.05 * 72)) -lt 4 -and
            [math]::Abs($shape.Top - (4.48 * 72)) -lt 4 -and
            $shape.Width -gt 280
        ) {
            $click = $shape.ActionSettings(1)
            $click.Action = 7
            $click.Hyperlink.Address = $wmvPath
        }
    }

    $pres.Save()
    $pres.SaveCopyAs($ppsxPath)
    $pres.Close()
} finally {
    if ($pp.Presentations.Count -gt 0) {
        $pp.Presentations | ForEach-Object { $_.Close() }
    }
    $pp.Quit()
}

Write-Output $pptxPath
Write-Output $ppsxPath
