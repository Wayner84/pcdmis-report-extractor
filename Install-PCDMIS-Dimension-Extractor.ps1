# Copyright (C) 2026 Townsend Precision Labs Ltd
# SPDX-License-Identifier: GPL-3.0-or-later

[CmdletBinding()]
param(
    [string]$TargetPath = 'C:\PCDMIS_Report_Fallback'
)

$ErrorActionPreference = 'Stop'
$sourcePath = Split-Path -Parent $PSCommandPath
$files = @(
    'Run-PCDMIS-Dimension-Extractor.bas',
    'Export-PCDMIS-Dimensions.bas',
    'Export-PCDMIS-Dimensions.ps1',
    'Required-PCDMIS-Variables.txt',
    'README.md',
    'LICENSE'
)
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'

New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
$installed = New-Object System.Collections.Generic.List[string]

foreach ($name in $files) {
    $source = Join-Path $sourcePath $name
    $destination = Join-Path $TargetPath $name
    if (-not (Test-Path -LiteralPath $source)) {
        throw "Package file is missing: $source"
    }
    if (Test-Path -LiteralPath $destination) {
        Copy-Item -LiteralPath $destination -Destination ($destination + '.backup-' + $timestamp) -Force
    }
    Copy-Item -LiteralPath $source -Destination $destination -Force
    $sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    $destinationHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
    if ($sourceHash -ne $destinationHash) {
        throw "Hash verification failed after installing $name"
    }
    $installed.Add("$name`t$destinationHash")
}

$logPath = Join-Path $TargetPath 'installation.log'
@(
    'PC-DMIS Dimension Extractor installation'
    "Installed=$((Get-Date).ToString('s'))"
    "Source=$sourcePath"
    "Target=$TargetPath"
    $installed
) | Set-Content -LiteralPath $logPath -Encoding UTF8

Write-Output "PC-DMIS Dimension Extractor installed to: $TargetPath"
Write-Output "Load this file in the PC-DMIS Basic Script Editor:"
Write-Output (Join-Path $TargetPath 'Run-PCDMIS-Dimension-Extractor.bas')
