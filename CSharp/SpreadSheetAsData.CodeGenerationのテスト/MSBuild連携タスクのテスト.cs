using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class MSBuild連携タスクのテスト
{
    [Fact]
    public void SpreadsheetAsData項目からobj配下へ生成コードを出力します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.Warnings.Should().BeEmpty();
        tested.GeneratedFilePaths.Should().ContainSingle();
        tested.SingleGeneratedFilePath.Should().StartWith(project.ObjDirectory);
        tested.SingleGeneratedFilePath.Should().EndWith(".g.cs");
        tested.SingleGeneratedSource
            .Should().Contain("public partial class BasicStructureBook : Workbook");
    }

    [Fact]
    public void プロジェクト直下の辞書を使用します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");
        project.AddProjectDictionaryFor(
            "BasicStructure.xlsx",
            """
            {
              "SalesData": "Sales"
            }
            """);

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource
            .Should().Contain("public SalesSheet Sales => new(this);");
    }

    [Fact]
    public void Excelファイルと同じディレクトリの辞書をプロジェクト直下の辞書より優先します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");
        project.AddProjectDictionaryFor(
            "BasicStructure.xlsx",
            """
            {
              "SalesData": "ProjectRootSales"
            }
            """);
        project.AddExcelDictionaryFor(
            @"Schemas\BasicStructure.xlsx",
            """
            {
              "SalesData": "SameDirectorySales"
            }
            """);

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.SingleGeneratedSource
            .Should().Contain("public SameDirectorySalesSheet SameDirectorySales => new(this);");
    }

    [Fact]
    public void 異なるディレクトリにある同名Excelファイルの生成ファイルは衝突しません()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var orderFilePath = project.AddBasicStructureExcel(@"Orders\Master.xlsx");
        var archiveFilePath = project.AddBasicStructureExcel(@"Archive\Master.xlsx");

        var tested = project.Generate(orderFilePath, archiveFilePath);

        tested.Succeeded.Should().BeTrue();
        tested.GeneratedFilePaths
            .Should().BeEquivalentTo(
                [
                    project.GeneratedFilePathFor(@"Orders\Master.xlsx"),
                    project.GeneratedFilePathFor(@"Archive\Master.xlsx")
                ]);
    }

    [Fact]
    public void 辞書を追加変更削除すると生成結果を更新します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public SalesDataSheet SalesData => new(this);");

        var dictionaryFilePath = project.AddProjectDictionaryFor(
            "BasicStructure.xlsx",
            """
            {
              "SalesData": "Sales"
            }
            """);

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public SalesSheet Sales => new(this);");

        File.WriteAllText(
            dictionaryFilePath,
            """
            {
              "SalesData": "Revenue"
            }
            """);

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public RevenueSheet Revenue => new(this);");

        File.Delete(dictionaryFilePath);

        project.Generate(excelFilePath).Succeeded.Should().BeTrue();
        project.GeneratedSourceFor(@"Schemas\BasicStructure.xlsx")
            .Should().Contain("public SalesDataSheet SalesData => new(this);");
    }

    [Fact]
    public void 不正な辞書JSONでは辞書ファイルを示して失敗します()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        var excelFilePath = project.AddBasicStructureExcel(@"Schemas\BasicStructure.xlsx");
        var dictionaryFilePath = project.AddExcelDictionaryFor(
            @"Schemas\BasicStructure.xlsx",
            "{");

        var tested = project.Generate(excelFilePath);

        tested.Succeeded.Should().BeFalse();
        tested.Errors.Should().ContainSingle();
        tested.Errors.Single().File.Should().Be(dictionaryFilePath);
    }

    [Fact]
    public void PowerShellからdotnet_msbuildでコード生成できます()
    {
        using var project = MSBuild連携テストプロジェクト.Create();
        project.AddBasicStructureExcel("BasicStructure.xlsx");
        var scriptFilePath = project.AddPowerShellGenerationSample();

        var tested = PowerShell実行結果.Run(
            scriptFilePath,
            project.DirectoryPath);

        tested.ExitCode.Should().Be(0, tested.Output);
        project.GeneratedSourceFor("BasicStructure.xlsx")
            .Should().Contain("public partial class BasicStructureBook : Workbook");
    }
}
