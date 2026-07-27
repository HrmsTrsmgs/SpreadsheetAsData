[CmdletBinding()]
param(
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$docfxConfig = [System.IO.Path]::Combine($repositoryRoot, "docs", "api", "csharp", "docfx.json")
$defaultOutputPath = [System.IO.Path]::Combine($repositoryRoot, "artifacts", "github-pages")
$pagesOutputPath = if ($OutputPath) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
}
else {
    $defaultOutputPath
}

function Invoke-Dotnet {
    dotnet @args
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet $($args -join ' ') failed with exit code $LASTEXITCODE."
    }
}

function New-Directory {
    param(
        [string]$Path
    )

    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

Push-Location $repositoryRoot
try {
    Invoke-Dotnet tool restore
    Invoke-Dotnet tool run docfx $docfxConfig

    if (Test-Path $pagesOutputPath) {
        Remove-Item -LiteralPath $pagesOutputPath -Recurse -Force
    }

    $csharpApiOutputPath = [System.IO.Path]::Combine($pagesOutputPath, "api", "csharp")
    New-Directory $csharpApiOutputPath

    Copy-Item `
        -Path ([System.IO.Path]::Combine($repositoryRoot, "docs", "api", "csharp", "_site", "*")) `
        -Destination $csharpApiOutputPath `
        -Recurse `
        -Force

    Set-Content `
        -Path (Join-Path $pagesOutputPath ".nojekyll") `
        -Value "" `
        -Encoding UTF8

    Set-Content `
        -Path (Join-Path $pagesOutputPath "index.html") `
        -Value @"
<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="refresh" content="0; url=api/csharp/">
  <title>SpreadsheetAsData</title>
</head>
<body>
  <p><a href="api/csharp/">SpreadsheetAsData C# API</a></p>
</body>
</html>
"@ `
        -Encoding UTF8

    Write-Host "GitHub Pages site was generated at $pagesOutputPath"
}
finally {
    Pop-Location
}
