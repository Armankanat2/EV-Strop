$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

function Resolve-PowerPointExe {
    $candidates = @(
        "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "POWERPNT.EXE not found in standard Office16 locations."
}

function Get-DeckPackageKind {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $contentTypes = $zip.Entries | Where-Object { $_.FullName -eq "[Content_Types].xml" }
        if (-not $contentTypes) {
            return "unknown"
        }

        $reader = New-Object System.IO.StreamReader($contentTypes.Open())
        try {
            $xml = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        if ($xml -match "presentationml\.slideshow\.main\+xml") {
            return "slideshow"
        }

        if ($xml -match "presentationml\.presentation\.main\+xml") {
            return "presentation"
        }

        return "unknown"
    }
    finally {
        if ($zip) {
            $zip.Dispose()
        }
    }
}

function Get-FreeSubstDrive {
    $preferred = @("P:", "Q:", "R:", "S:", "T:", "U:", "V:", "W:", "X:", "Y:", "Z:")
    $used = (Get-PSDrive -PSProvider FileSystem).Name | ForEach-Object { "${_}:" }

    foreach ($drive in $preferred) {
        if ($used -notcontains $drive) {
            return $drive
        }
    }

    throw "No free drive letter available for SUBST alias."
}

function New-AsciiPathAlias {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $parentDir = Split-Path -Path $Path -Parent
    $leafName = Split-Path -Path $Path -Leaf
    $resolvedParent = if ([string]::IsNullOrWhiteSpace($parentDir)) {
        (Resolve-Path -LiteralPath ".").Path
    }
    else {
        (Resolve-Path -LiteralPath $parentDir).Path
    }

    $resolvedPath = Join-Path $resolvedParent $leafName
    if ($resolvedPath -cmatch '^[\u0000-\u007F]+$') {
        return [pscustomobject]@{
            Path  = $resolvedPath
            Drive = $null
        }
    }

    $drive = Get-FreeSubstDrive
    & subst $drive $resolvedParent | Out-Null

    $aliasPath = Join-Path $drive $leafName
    return [pscustomobject]@{
        Path  = $aliasPath
        Drive = $drive
    }
}

function Remove-AsciiPathAlias {
    param(
        [Parameter()]
        $Alias
    )

    if ($null -ne $Alias -and $Alias.Drive) {
        & subst $Alias.Drive /d | Out-Null
    }
}

function Convert-ToStablePpsx {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SourcePath,

        [string] $OutputPath
    )

    $resolvedSource = (Resolve-Path -LiteralPath $SourcePath).Path
    if (-not $OutputPath) {
        $sourceItem = Get-Item -LiteralPath $resolvedSource
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sourceItem.Name)
        $OutputPath = Join-Path $sourceItem.DirectoryName ($baseName + "_stable-open.ppsx")
    }

    Export-StablePpsx -SourcePath $resolvedSource -OutputPath $OutputPath
}

function Export-StablePpsx {
    param(
        [Parameter(Mandatory = $true)]
        [string] $SourcePath,

        [Parameter(Mandatory = $true)]
        [string] $OutputPath
    )

    $resolvedSource = (Resolve-Path -LiteralPath $SourcePath).Path
    $sourceParent = Split-Path -Path $resolvedSource -Parent
    $outputParent = Split-Path -Path $OutputPath -Parent
    $resolvedOutputParent = (Resolve-Path -LiteralPath $outputParent).Path
    $resolvedOutput = Join-Path $resolvedOutputParent (Split-Path -Path $OutputPath -Leaf)

    $sourceAlias = $null
    $outputAlias = $null
    $pp = $null
    $presentation = $null

    try {
        if ($sourceParent -eq $resolvedOutputParent) {
            $sharedAlias = New-AsciiPathAlias -Path $resolvedSource
            $sourceAlias = $sharedAlias
            $outputAlias = [pscustomobject]@{
                Path  = Join-Path (Split-Path -Path $sharedAlias.Path -Parent) (Split-Path -Path $resolvedOutput -Leaf)
                Drive = $sharedAlias.Drive
            }
        }
        else {
            $sourceAlias = New-AsciiPathAlias -Path $resolvedSource
            $outputAlias = New-AsciiPathAlias -Path $resolvedOutput
        }

        if (Test-Path -LiteralPath $resolvedOutput) {
            Remove-Item -LiteralPath $resolvedOutput -Force
        }

        $pp = New-Object -ComObject PowerPoint.Application
        $pp.Visible = -1

        $presentation = $pp.Presentations.Open($sourceAlias.Path, $false, $false, $false)
        $presentation.SaveAs(
            $outputAlias.Path,
            [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLShow
        )
    }
    finally {
        if ($presentation) {
            $presentation.Close()
        }
        if ($pp) {
            $pp.Quit()
        }
        if ($sourceAlias -and $outputAlias -and $sourceAlias.Drive -and $sourceAlias.Drive -eq $outputAlias.Drive) {
            Remove-AsciiPathAlias -Alias $sourceAlias
        }
        else {
            Remove-AsciiPathAlias -Alias $sourceAlias
            Remove-AsciiPathAlias -Alias $outputAlias
        }
    }

    return $resolvedOutput
}

function Open-PowerPointFile {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    $powerPointExe = Resolve-PowerPointExe
    $alias = New-AsciiPathAlias -Path $resolvedPath
    $extension = [System.IO.Path]::GetExtension($resolvedPath).ToLowerInvariant()
    $launchArgs = if ($extension -eq ".ppsx") {
        @("/s", $alias.Path)
    }
    else {
        @($alias.Path)
    }

    # Keep the SUBST alias alive after launch so PowerPoint can finish opening the file.
    Start-Process -FilePath $powerPointExe -ArgumentList $launchArgs

    return $resolvedPath
}
