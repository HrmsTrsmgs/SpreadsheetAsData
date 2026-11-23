[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$failed = $false

function Write-CheckResult {
    param(
        [string]$Name,
        [bool]$Succeeded,
        [string]$Message
    )

    if ($Succeeded) {
        Write-Host "[OK] $Name - $Message"
    }
    else {
        Write-Host "[NG] $Name - $Message"
        $script:failed = $true
    }
}

function Get-CommandText {
    param([string]$Name)

    $command = Get-Command $Name -ErrorAction SilentlyContinue

    if ($null -eq $command) {
        return $null
    }

    return $command.Source
}

$pwshPath = Get-CommandText "pwsh"
Write-CheckResult `
    "PowerShell 7" `
    ($null -ne $pwshPath) `
    ($(if ($null -eq $pwshPath) {
        "pwsh が見つかりません。winget install --id Microsoft.PowerShell -e などでインストールしてください。"
    }
    else {
        "pwsh: $pwshPath"
    }))

Write-CheckResult `
    "Current shell" `
    ($PSVersionTable.PSVersion.Major -ge 7) `
    "current: $($PSVersionTable.PSVersion)。このリポジトリのPowerShellスクリプトは pwsh -NoProfile -File <script> で実行してください。"

$dotnetPath = Get-CommandText "dotnet"
Write-CheckResult `
    ".NET SDK" `
    ($null -ne $dotnetPath) `
    ($(if ($null -eq $dotnetPath) {
        "dotnet が見つかりません。.NET 11 SDKをインストールしてください。"
    }
    else {
        "dotnet: $dotnetPath"
    }))

if ($null -ne $dotnetPath) {
    $dotnetSdks = dotnet --list-sdks
    $dotnet11Sdks = @(
        $dotnetSdks |
            Where-Object { $_ -match "^11\." }
    )

    Write-CheckResult `
        ".NET 11 SDK" `
        ($dotnet11Sdks.Count -gt 0) `
        ($(if ($dotnet11Sdks.Count -gt 0) {
            $dotnet11Sdks -join ", "
        }
        else {
            ".NET 11 SDKが見つかりません。"
        }))
}

$gitPath = Get-CommandText "git"
Write-CheckResult `
    "Git" `
    ($null -ne $gitPath) `
    ($(if ($null -eq $gitPath) {
        "git が見つかりません。"
    }
    else {
        "git: $gitPath"
    }))

if ($failed) {
    exit 1
}
