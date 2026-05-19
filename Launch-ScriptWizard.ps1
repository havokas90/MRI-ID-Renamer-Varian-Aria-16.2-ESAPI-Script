$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptPath = Join-Path $projectRoot 'InstallReady\EclipseImageRenamer.cs'
$scriptWizardPath = 'C:\Program Files\Varian\RTM\17.0\esapi\ScriptWizard.exe'

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Install-ready script was not found: $scriptPath"
}

if (-not (Test-Path -LiteralPath $scriptWizardPath)) {
    throw "Varian Script Wizard was not found: $scriptWizardPath"
}

Start-Process -FilePath $scriptWizardPath -ArgumentList "`"$scriptPath`""
