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
                IntegratedExcelFilePath,
                options => options.Namespace = "Generated.Custom")
            .GeneratedTypes;

        generatedTypes
            .Should()
            .NotBeEmpty();

        generatedTypes
            .Should()
            .OnlyContain(it => it.NamespaceName == "Generated.Custom");
    }

    [Fact(
        Skip =
            "基本的なExcelファイルから生成したソースがRoslynでコンパイルできる処理を実装するときに解除する。")]
    public void 基本的なExcelファイルからコンパイルできるソースを生成します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(IntegratedExcelFilePath)
            .Should()
            .NotBeNull();
    }

    [Fact(
        Skip =
            "同じExcelファイルと同じ設定から決定的な生成結果を返す処理を実装するときに解除する。")]
    public void 同じExcelファイルと設定から同じ生成結果を返します()
    {
        GeneratedCodeInspection
            .GenerateSources(IntegratedExcelFilePath)
            .Should()
            .Equal(GeneratedCodeInspection.GenerateSources(IntegratedExcelFilePath));
    }

    [Fact(
        Skip =
            "正常なExcelファイルでエラー診断を返さない処理を実装するときに解除する。")]
    public void 正常なExcelファイルではエラー診断を返しません()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(IntegratedExcelFilePath)
            .Should()
            .BeEmpty();
    }
}


