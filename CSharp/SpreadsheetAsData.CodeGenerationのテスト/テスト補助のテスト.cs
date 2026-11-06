using FluentAssertions;
using Marimo.SpreadsheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadsheetAsData.CodeGeneration.Test;

public sealed class テスト補助のテスト
{
    [Fact]
    public async Task PowerShell実行結果は大量の標準出力と標準エラーを回収して終了します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        Directory.CreateDirectory(project.DirectoryPath);
        var processIdPath = Path.Combine(project.DirectoryPath, "pwsh.pid");
        var scriptFilePath = Path.Combine(project.DirectoryPath, "LargeOutput.ps1");
        File.WriteAllText(
            scriptFilePath,
            $$"""
            Set-Content -LiteralPath '{{processIdPath.Replace("'", "''")}}' -Value $PID
            [Console]::Error.Write('E' * 262144)
            [Console]::Out.WriteLine('completed')
            """);

        var run = PowerShell実行結果.RunAsync(scriptFilePath, project.DirectoryPath);
        try
        {
            var tested = await run.WaitAsync(TimeSpan.FromSeconds(15), TestContext.Current.CancellationToken);

            tested.ExitCode.Should().Be(0);
            tested.Output.Should().Contain("completed");
            tested.Output.Should().Contain(new string('E', 262144));
        }
        finally
        {
            if (!run.IsCompleted && File.Exists(processIdPath))
            {
                try
                {
                    using var process = System.Diagnostics.Process.GetProcessById(
                        int.Parse(File.ReadAllText(processIdPath)));
                    if (!process.HasExited)
                    {
                        process.Kill(entireProcessTree: true);
                    }
                }
                catch (ArgumentException)
                {
                    // The child exited between the timeout and cleanup.
                }
                catch (InvalidOperationException)
                {
                    // The child exited while it was being stopped.
                }
            }
        }
    }

    [Fact]
    public void TypeNamesは生成ソースに含まれる型名を返します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                """
                namespace Generated;

                public class BasicStructureBook
                {
                }

                public class SalesDataSheet
                {
                }
                """)
            .TypeNames
            .Should().BeEquivalentTo("BasicStructureBook", "SalesDataSheet");
    }

    [Fact]
    public void GeneratedTypeは指定した名前の生成型を返します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                """
                namespace Generated;

                public class BasicStructureBook
                {
                }

                public class SalesDetailTable
                {
                }
                """)
            .GeneratedType("SalesDetailTable")
            .Name
            .Should().Be("SalesDetailTable");
    }

    [Fact]
    public void Nameは生成型の名前を返します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                """
                namespace Generated;

                public class SalesDetailTable
                {
                }
                """)
            .GeneratedType("SalesDetailTable")
            .Name
            .Should().Be("SalesDetailTable");
    }

    [Fact]
    public void NamespaceNameは指定した生成型の名前空間を返します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                """
                namespace Generated.Custom;

                public class SalesDataSheet
                {
                }
                """)
            .GeneratedType("SalesDataSheet")
            .NamespaceName
            .Should().Be("Generated.Custom");
    }

    [Fact]
    public void ToStringは生成型の名前空間付き表示名を返します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                """
                namespace Generated;

                public class SalesDataSheet
                {
                }
                """)
            .GeneratedType("SalesDataSheet")
            .ToString().Should().Be("SalesDataSheet: Generated");
    }

    [Fact]
    public void PropertyNamesは指定した型に含まれるプロパティ名を返します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                """
                namespace Generated;

                public class SalesDetail
                {
                    public int CustomerId { get; set; }
                    public double Amount { get; set; }
                }
                """)
            .GeneratedType("SalesDetail")
            .PropertyNames
            .Should().BeEquivalentTo("CustomerId", "Amount");
    }
}
