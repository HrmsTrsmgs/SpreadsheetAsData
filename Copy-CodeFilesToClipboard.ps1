function Copy-CodeFilesToClipboard {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [string]$Path = ".",

        [Parameter(Mandatory, Position = 1)]
        [string[]]$Extensions,

        [string[]]$ExcludeDirectories = @(
            ".git",
            "node_modules",
            ".next",
            "dist",
            "build",
            "out",
            "coverage"
        ),

        [string[]]$ExcludeFiles = @(
            "package-lock.json",
            "pnpm-lock.yaml",
            "yarn.lock"
        ),

        [switch]$PassThru
    )

    $root = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $root = $root.TrimEnd([char[]]"\/")

    $normalizedExtensions = $Extensions |
        ForEach-Object {
            $extension = $_.Trim()

            if (-not $extension.StartsWith(".")) {
                $extension = ".$extension"
            }

            $extension.ToLowerInvariant()
        } |
        Select-Object -Unique

    $files = Get-ChildItem `
        -LiteralPath $root `
        -Recurse `
        -File `
        -Force |
        Where-Object {
            if ($normalizedExtensions -notcontains $_.Extension.ToLowerInvariant()) {
                return $false
            }

            if ($ExcludeFiles -contains $_.Name) {
                return $false
            }

            $relativePath = $_.FullName.Substring($root.Length).TrimStart([char[]]"\/")

            $relativeDirectory =
                [System.IO.Path]::GetDirectoryName($relativePath)

            if (-not [string]::IsNullOrWhiteSpace($relativeDirectory)) {
                foreach ($segment in $relativeDirectory -split "[\\/]") {
                    if ($ExcludeDirectories -contains $segment) {
                        return $false
                    }
                }
            }

            return $true
        } |
        Sort-Object FullName

    if ($files.Count -eq 0) {
        throw "指定された拡張子のファイルが見つかりませんでした: $($normalizedExtensions -join ', ')"
    }

    $builder = [System.Text.StringBuilder]::new()

    foreach ($file in $files) {
        $relativePath = $file.FullName.Substring($root.Length).TrimStart([char[]]"\/").Replace("\", "/")

        try {
            $content = [System.IO.File]::ReadAllText($file.FullName)
        }
        catch {
            Write-Warning "読み込みに失敗しました: $relativePath"
            continue
        }

        [void]$builder.AppendLine(
            "===== BEGIN FILE: $relativePath ====="
        )
        [void]$builder.AppendLine($content)
        [void]$builder.AppendLine(
            "===== END FILE: $relativePath ====="
        )
        [void]$builder.AppendLine()
    }

    $result = $builder.ToString()

    Set-Clipboard -Value $result

    Write-Host (
        "{0}ファイル、{1:N0}文字をクリップボードへコピーしました。" -f
        $files.Count,
        $result.Length
    )

    if ($PassThru) {
        $result
    }
}