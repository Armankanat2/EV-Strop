$ErrorActionPreference = "Stop"

function Set-Text {
    param(
        [Parameter(Mandatory = $true)] $Shape,
        [Parameter(Mandatory = $true)] [string] $Text,
        [int] $FontSize = 0,
        [bool] $Bold = $false
    )

    $Shape.TextFrame.TextRange.Text = $Text
    if ($FontSize -gt 0) {
        $Shape.TextFrame.TextRange.Font.Size = $FontSize
    }
    $Shape.TextFrame.TextRange.Font.Bold = [int] $Bold
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

function Add-VideoToSlide {
    param(
        [Parameter(Mandatory = $true)] $Slide,
        [Parameter(Mandatory = $true)] [string] $VideoPath,
        [Parameter(Mandatory = $true)] [single] $Left,
        [Parameter(Mandatory = $true)] [single] $Top,
        [Parameter(Mandatory = $true)] [single] $Width,
        [Parameter(Mandatory = $true)] [single] $Height,
        [string] $Name = "VIDEO"
    )

    $media = $Slide.Shapes.AddMediaObject2($VideoPath, $false, $true, $Left, $Top, $Width, $Height)
    $media.Name = $Name

    $playSettings = $media.AnimationSettings.PlaySettings
    $playSettings.PlayOnEntry = -1
    $playSettings.HideWhileNotPlaying = 0
    $playSettings.LoopUntilStopped = 0
    $playSettings.PauseAnimation = 0
    $playSettings.RewindMovie = -1
    $playSettings.StopAfterSlides = 1

    return $media
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $scriptDir "S001-S040_live_preview_working_2026-06-29_v52.pptx"
$outputPath = Join-Path $scriptDir "S001-S040_live_preview_working_2026-06-29_v53.pptx"
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptDir)))
$videoSourcePath = Join-Path $repoRoot "assets\course-media\module-01-stropovka-gruzov\animations\angle-between-branches-master-v5_2026-06-26\angle-between-branches_master-v5_2026-06-26.mp4"
$videoPath = Join-Path $scriptDir "angle-between-branches_master-v5_2026-06-26.wmv"
$ffmpegPath = "C:\Program Files (x86)\ClipGrab\ffmpeg.exe"

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Source presentation not found: $sourcePath"
}

if (-not (Test-Path -LiteralPath $videoSourcePath)) {
    throw "Source video file not found: $videoSourcePath"
}

$videoSourcePath = (Resolve-Path -LiteralPath $videoSourcePath).Path

if (-not (Test-Path -LiteralPath $ffmpegPath)) {
    throw "FFmpeg not found: $ffmpegPath"
}

if (Test-Path -LiteralPath $outputPath) {
    Remove-Item -LiteralPath $outputPath -Force
}

Copy-Item -LiteralPath $sourcePath -Destination $outputPath -Force

if (-not (Test-Path -LiteralPath $videoPath)) {
    & $ffmpegPath -y -i $videoSourcePath -c:v wmv2 -b:v 1800k -c:a wmav2 -b:a 128k $videoPath | Out-Null
}

if (-not (Test-Path -LiteralPath $videoPath)) {
    throw "Converted WMV file not found: $videoPath"
}

$pp = $null
$presentation = $null

try {
    $pp = New-Object -ComObject PowerPoint.Application
    $pp.Visible = -1
    $presentation = $pp.Presentations.Open($outputPath, $false, $false, $false)

    $presentation.Slides.Item(74).Duplicate() | Out-Null
    $presentation.Slides.Item(75).MoveTo(75)

    $s021 = $presentation.Slides.Item(73)
    $s022 = $presentation.Slides.Item(74)
    $s023 = $presentation.Slides.Item(75)
    $s024 = $presentation.Slides.Item(76)
    $s025 = $presentation.Slides.Item(77)
    $s026 = $presentation.Slides.Item(78)
    $s027 = $presentation.Slides.Item(79)
    $s028 = $presentation.Slides.Item(80)

    Set-Text -Shape $s022.Shapes.Item(2) -Text "Защита стропа и груза" -FontSize 24 -Bold $true
    Set-Text -Shape $s022.Shapes.Item(3) -Text "Острые кромки и хрупкий груз требуют защиты в точке контакта" -FontSize 18
    Set-Text -Shape $s022.Shapes.Item(4) -Text "S022" -FontSize 14 -Bold $true
    Set-Text -Shape $s022.Shapes.Item(6) -Text "Что учитывать" -FontSize 18 -Bold $true
    Set-Text -Shape $s022.Shapes.Item(7) -Text "Защиту подбирают под кромку, материал груза и схему строповки." -FontSize 12
    Set-Text -Shape $s022.Shapes.Item(9) -Text "Видео: нагрузка при увеличении угла" -FontSize 18 -Bold $true
    Set-Text -Shape $s022.Shapes.Item(10) -Text "• Острые кромки груза могут повредить текстильный строп.`r• Канатный строп может повредить хрупкий груз в месте контакта.`r• Между стропом и грузом нужна защита, если есть риск пореза, смятия или скола." -FontSize 16
    $s022.Shapes.Item(13).Delete()
    Add-VideoToSlide -Slide $s022 -VideoPath $videoPath -Left ([single] 521) -Top ([single] 176) -Width ([single] 350) -Height ([single] 248) -Name "S022_ANGLE_VIDEO" | Out-Null

    Set-Text -Shape $s023.Shapes.Item(2) -Text "Протектор между стропом и грузом" -FontSize 24 -Bold $true
    Set-Text -Shape $s023.Shapes.Item(3) -Text "Защитная прокладка ставится в зоне контакта стропа с кромкой или чувствительной поверхностью" -FontSize 18
    Set-Text -Shape $s023.Shapes.Item(4) -Text "S023" -FontSize 14 -Bold $true
    Set-Text -Shape $s023.Shapes.Item(6) -Text "Когда нужен протектор" -FontSize 18 -Bold $true
    Set-Text -Shape $s023.Shapes.Item(7) -Text "Нужны примеры: защита текстильного стропа и защита хрупкого груза." -FontSize 12
    Set-Text -Shape $s023.Shapes.Item(9) -Text "Визуал / слайд-шоу" -FontSize 18 -Bold $true
    Set-Text -Shape $s023.Shapes.Item(10) -Text "• При острых кромках и ребрах.`r• При риске перетирания или пореза текстильного стропа.`r• При контакте канатного стропа с хрупким грузом.`r• Протектор должен закрывать место контакта на всем этапе подъема." -FontSize 16
    Set-Text -Shape $s023.Shapes.Item(13) -Text "Нужны 2-3 кадра или короткое слайд-шоу: протектор на острой кромке; прокладка под канатный строп на хрупком грузе; крупный план правильной установки." -FontSize 16

    Set-Text -Shape $s024.Shapes.Item(4) -Text "S024" -FontSize 14 -Bold $true
    if ($s024.Shapes.Count -ge 25) {
        $s024.Shapes.Item(25).Delete()
    }

    Set-Text -Shape $s025.Shapes.Item(4) -Text "S025" -FontSize 14 -Bold $true

    Set-Text -Shape $s026.Shapes.Item(4) -Text "S026" -FontSize 14 -Bold $true
    Set-Text -Shape $s026.Shapes.Item(23) -Text "• Связать правила темы с реальными примерами схем. `r• Подвести слушателя к проверке понимания: где ошибка и почему она опасна. `r• Сделать мост к тесту S027." -FontSize 16

    Set-Text -Shape $s027.Shapes.Item(4) -Text "S027" -FontSize 14 -Bold $true
    Set-Text -Shape $s028.Shapes.Item(4) -Text "S028" -FontSize 14 -Bold $true

    Set-SlideJump -Shape $s022.Shapes.Item(11) -TargetSlide $s021
    Set-SlideJump -Shape $s022.Shapes.Item(12) -TargetSlide $s023

    Set-SlideJump -Shape $s023.Shapes.Item(11) -TargetSlide $s022
    Set-SlideJump -Shape $s023.Shapes.Item(12) -TargetSlide $s024

    Set-SlideJump -Shape $s024.Shapes.Item(23) -TargetSlide $s023
    Set-SlideJump -Shape $s024.Shapes.Item(24) -TargetSlide $s025

    Set-SlideJump -Shape $s025.Shapes.Item(12) -TargetSlide $s024
    Set-SlideJump -Shape $s025.Shapes.Item(13) -TargetSlide $s026

    Set-SlideJump -Shape $s026.Shapes.Item(24) -TargetSlide $s025
    Set-SlideJump -Shape $s026.Shapes.Item(25) -TargetSlide $s027

    Set-SlideJump -Shape $s027.Shapes.Item(13) -TargetSlide $s026
    Set-SlideJump -Shape $s027.Shapes.Item(14) -TargetSlide $s028

    Set-SlideJump -Shape $s028.Shapes.Item(24) -TargetSlide $s027

    $presentation.Save()
}
finally {
    if ($presentation) {
        $presentation.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
    }
    if ($pp) {
        $pp.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
    }
}

Write-Output $outputPath
