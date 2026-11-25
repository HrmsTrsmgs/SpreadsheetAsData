[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repositoryRoot "CSharp\SpreadsheetAsData.slnx"

. (Join-Path $PSScriptRoot "Invoke-Dotnet.ps1")

& (Join-Path $PSScriptRoot "Test-DevelopmentEnvironment.ps1")

Push-Location $repositoryRoot
try {
    Invoke-Dotnet tool restore
    Invoke-Dotnet restore $solutionPath
}
finally {
    Pop-Location
}

Write-Host "SpreadsheetAsData development environment is ready."
