[CmdletBinding()]
param(
    [switch]$BuildOnly
)

$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$docfxConfig = Join-Path $repositoryRoot "docs\api\csharp\docfx.json"

. (Join-Path $PSScriptRoot "Invoke-Dotnet.ps1")

Push-Location $repositoryRoot
try {
    Invoke-Dotnet tool restore

    if ($BuildOnly) {
        Invoke-Dotnet tool run docfx $docfxConfig
        return
    }

    Write-Host "Serving SpreadsheetAsData C# API docs. Press Ctrl+C to stop."
    Invoke-Dotnet tool run docfx $docfxConfig --serve
}
finally {
    Pop-Location
}
