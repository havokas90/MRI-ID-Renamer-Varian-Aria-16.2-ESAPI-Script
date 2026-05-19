param(
    [switch]$CheckOnly
)

$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $projectRoot 'EclipseImageRenamer.cs'
$installReadyPath = Join-Path $projectRoot 'InstallReady\EclipseImageRenamer.cs'

if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Source file not found: $sourcePath"
}

if (-not (Test-Path -LiteralPath $installReadyPath)) {
    throw "InstallReady file not found: $installReadyPath"
}

$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$installReadyHash = (Get-FileHash -LiteralPath $installReadyPath -Algorithm SHA256).Hash

if ($sourceHash -eq $installReadyHash) {
    Write-Output 'InstallReady/EclipseImageRenamer.cs is in sync.'
    exit 0
}

if ($CheckOnly) {
    Write-Error 'InstallReady/EclipseImageRenamer.cs is out of sync. Run .\Sync-InstallReady.ps1 after editing EclipseImageRenamer.cs.'
    exit 1
}

Copy-Item -LiteralPath $sourcePath -Destination $installReadyPath -Force
Write-Output 'Synced EclipseImageRenamer.cs to InstallReady/EclipseImageRenamer.cs.'
