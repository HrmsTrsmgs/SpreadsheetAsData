[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repositoryRoot "CSharp\SpreadSheetAsData.slnx"

function Invoke-Dotnet {
    dotnet @args

    if ($LASTEXITCODE -ne 0) {
        throw "dotnet $($args -join ' ') failed with exit code $LASTEXITCODE."
    }
}

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
