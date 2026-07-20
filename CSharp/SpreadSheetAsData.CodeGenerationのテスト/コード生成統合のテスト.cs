using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成統合のテスト
{
    const string IntegratedExcelFilePath = @"TestData\コード生成\統合.xlsx";

    [Fact]
    public void 指定した名前空間へすべての型を生成します()
    {
        var generatedTypes = GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    IntegratedExcelFilePath,
                    options => options.Namespace = "Generated.Custom"))
            .GeneratedTypes;

        generatedTypes
            .Should().NotBeEmpty();

        generatedTypes
            .Should().OnlyContain(it => it.NamespaceName == "Generated.Custom");
    }

    [Fact]
    public void 基本的なExcelファイルからコンパイルできるソースを生成します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    IntegratedExcelFilePath))
            .Should().NotBeNull();
    }

    [Fact]
    public void 同じExcelファイルと設定から同じ生成結果を返します()
    {
        GeneratedCodeInspection
            .GenerateSources(IntegratedExcelFilePath)
            .Should().Equal(GeneratedCodeInspection.GenerateSources(
                    IntegratedExcelFilePath));
    }

    [Fact]
    public void 正常なExcelファイルではエラー診断を返しません()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(IntegratedExcelFilePath)
            .Should().BeEmpty();
    }
}


