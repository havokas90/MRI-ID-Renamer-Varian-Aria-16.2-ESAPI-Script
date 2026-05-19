$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$previewProject = Join-Path $projectRoot 'LocalPreview\ImageIdRenamer.LocalPreview.csproj'

if (-not (Test-Path -LiteralPath $previewProject)) {
    throw "Local preview project was not found: $previewProject"
}

dotnet run --project $previewProject
