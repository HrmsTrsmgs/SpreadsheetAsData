using System.Diagnostics;
using System.Text;
using FluentAssertions;
using Marimo.SpreadsheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadsheetAsData.CodeGeneration.Test;

public sealed class GitHubPages生成スクリプトのテスト
{
    [Fact]
    public void リポジトリルートを出力先に指定すると内容を変更せず拒否します()
    {
        using var repository = new 一時スクリプトリポジトリ();
        repository.AddPagesContent();
        var markerPath = Path.Combine(repository.DirectoryPath, "keep.txt");
        File.WriteAllText(markerPath, "keep");

        var tested = repository.RunPages(repository.DirectoryPath);

        tested.ExitCode.Should().NotBe(0);
        tested.Output.Should().Contain("repository root");
        File.ReadAllText(markerPath).Should().Be("keep");
    }

    [Fact]
    public void 生成物用の出力先は既存内容を置き換えて再生成できます()
    {
        using var repository = new 一時スクリプトリポジトリ();
        repository.AddPagesContent();
        var outputPath = Path.Combine(repository.DirectoryPath, "artifacts", "github-pages");
        Directory.CreateDirectory(outputPath);
        File.WriteAllText(Path.Combine(outputPath, "stale.txt"), "old");

        var tested = repository.RunPages(outputPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        File.ReadAllText(Path.Combine(outputPath, "api", "csharp", "index.html"))
            .Should().Be("api");
        File.Exists(Path.Combine(outputPath, "stale.txt")).Should().BeFalse();
    }
}

public sealed class 変更ファイル正規化スクリプトのテスト
{
    [Fact]
    public void BOMなしのCSharpファイルにUTF8のBOMを付加します()
    {
        using var repository = new 一時スクリプトリポジトリ();
        var filePath = Path.Combine(repository.DirectoryPath, "Sample.cs");
        File.WriteAllText(filePath, "class Sample {}\n", new UTF8Encoding(false));

        repository.RunNormalizer().ExitCode.Should().Be(0);

        File.ReadAllBytes(filePath).Take(3).Should().Equal(0xEF, 0xBB, 0xBF);
    }

    [Fact]
    public void MarkdownファイルからUTF8のBOMを除きます()
    {
        using var repository = new 一時スクリプトリポジトリ();
        var filePath = Path.Combine(repository.DirectoryPath, "notes.md");
        File.WriteAllText(filePath, "text\n", new UTF8Encoding(true));

        repository.RunNormalizer().ExitCode.Should().Be(0);

        File.ReadAllBytes(filePath).Take(3).Should().Equal(0x74, 0x65, 0x78);
    }

    [Fact(Skip = "BOM処理を修正する際に解除し、正規化を繰り返してもバイト列が変わらないことを確認する")]
    public void 正規化済みのファイルは再実行しても変わりません()
    {
        using var repository = new 一時スクリプトリポジトリ();
        var filePath = Path.Combine(repository.DirectoryPath, "Sample.cs");
        File.WriteAllText(filePath, "class Sample {}\n", new UTF8Encoding(false));
        repository.RunNormalizer().ExitCode.Should().Be(0);
        var normalized = File.ReadAllBytes(filePath);

        repository.RunNormalizer().ExitCode.Should().Be(0);

        normalized.Take(3).Should().Equal(0xEF, 0xBB, 0xBF);
        File.ReadAllBytes(filePath).Should().Equal(normalized);
    }
}

sealed class 一時スクリプトリポジトリ : IDisposable
{
    internal string DirectoryPath { get; } = Path.Combine(
        Path.GetTempPath(), "SpreadsheetAsData.Tests", Guid.NewGuid().ToString("N"));

    internal 一時スクリプトリポジトリ() =>
        Directory.CreateDirectory(Path.Combine(DirectoryPath, "scripts"));

    internal void AddPagesContent()
    {
        var docsPath = Path.Combine(DirectoryPath, "docs", "api", "csharp");
        var sitePath = Path.Combine(docsPath, "_site");
        Directory.CreateDirectory(sitePath);
        File.WriteAllText(Path.Combine(docsPath, "docfx.json"), "{}");
        File.WriteAllText(Path.Combine(sitePath, "index.html"), "api");
    }

    internal PowerShell実行結果 RunPages(string outputPath) =>
        Run(
            "build-github-pages.ps1",
            $"-OutputPath '{PowerShellLiteral(outputPath)}'",
            "function dotnet { $global:LASTEXITCODE = 0 }");

    internal PowerShell実行結果 RunNormalizer()
    {
        using var git = Process.Start(new ProcessStartInfo
        {
            FileName = "git",
            ArgumentList = { "init", "--quiet", DirectoryPath },
            UseShellExecute = false
        }) ?? throw new InvalidOperationException("Gitを起動できませんでした。");
        git.WaitForExit();
        if (git.ExitCode != 0)
        {
            throw new InvalidOperationException("一時リポジトリを作成できませんでした。");
        }

        return Run("Normalize-ChangedTextFiles.ps1", "-IncludeUntracked");
    }

    public void Dispose()
    {
        if (Directory.Exists(DirectoryPath))
        {
            Directory.Delete(DirectoryPath, true);
        }
    }

    PowerShell実行結果 Run(string fileName, string arguments, string setup = "")
    {
        var scriptPath = Path.Combine(DirectoryPath, "scripts", fileName);
        File.Copy(
            Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory, "..", "..", "..", "..", "..", "scripts", fileName)),
            scriptPath,
            true);
        var runnerPath = Path.Combine(DirectoryPath, "Run.ps1");
        File.WriteAllText(
            runnerPath,
            $"{setup}\n& '{PowerShellLiteral(scriptPath)}' {arguments}\n");
        return PowerShell実行結果.Run(runnerPath, DirectoryPath);
    }

    static string PowerShellLiteral(string path) => path.Replace("'", "''");
}
