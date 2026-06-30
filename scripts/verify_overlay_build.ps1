param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Resolve-FullPath {
    param([string]$path)

    if ([System.IO.Path]::IsPathRooted($path)) {
        return $path
    }

    return (Join-Path $PWD $path)
}

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

$src = Resolve-FullPath $SourcePath
$dst = Resolve-FullPath $OutputPath

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
