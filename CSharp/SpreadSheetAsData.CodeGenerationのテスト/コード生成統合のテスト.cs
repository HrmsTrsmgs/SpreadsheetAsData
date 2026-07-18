using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成統合のテスト
{
    const string TestFilePath = @"TestData\コード生成\統合.xlsx";

    [Fact(
        Skip =
            "指定した名前空間へすべての生成型を出力する処理を実装するときに解除する。")]
    public void 指定した名前空間へすべての型を生成します()
    {
        CodeGenerationSpec
            .GenerateSources(
                TestFilePath,
                options => options.Namespace = "Generated.Custom")
            .TypeDeclarations()
            .Should()
            .OnlyContain(it => it.Parent != null
                && it.Parent.ToString().Contains("Generated.Custom"));
    }

    [Fact(
        Skip =
            "有効なExcelファイルから生成したすべてのソースがRoslynでコンパイルできる処理を実装するときに解除する。")]
    public void 有効なExcelファイルから生成したすべてのソースはコンパイルできます()
    {
        CodeGenerationSpec
            .CompileGeneratedAssembly(TestFilePath)
            .Should()
            .NotBeNull();
    }

    [Fact(
        Skip =
            "同じExcelファイルと同じ設定から決定的な生成結果を返す処理を実装するときに解除する。")]
    public void 同じExcelファイルと設定から同じ生成結果を返します()
    {
        CodeGenerationSpec
            .GenerateSources(TestFilePath)
            .Should()
            .Equal(CodeGenerationSpec.GenerateSources(TestFilePath));
    }

    [Fact(
        Skip =
            "正常なExcelファイルでエラー診断を返さない処理を実装するときに解除する。")]
    public void 正常なExcelファイルではエラー診断を返しません()
    {
        CodeGenerationSpec
            .GenerateDiagnostics(TestFilePath)
            .Should()
            .BeEmpty();
    }
}


