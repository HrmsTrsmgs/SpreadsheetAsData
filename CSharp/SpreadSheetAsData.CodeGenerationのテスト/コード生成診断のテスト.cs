using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成診断のテスト
{
    const string ColumnNameCollisionExcelFilePath = @"TestData\コード生成\列名衝突.xlsx";
    const string BookMemberNameCollisionExcelFilePath = @"TestData\コード生成\Bookメンバー名衝突.xlsx";
    const string InvalidNameExcelFilePath = @"TestData\コード生成\無効名.xlsx";

    [Fact(
        Skip =
            "自動変換後に同じ列プロパティ名となる場合のエラー診断を実装するときに解除する。")]
    public void 自動変換後に同じ列プロパティ名となる場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(ColumnNameCollisionExcelFilePath)
            .Should()
            .ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "CustomerId",
                    ["customer_id", "customer-id"]));
    }

    [Fact(
        Skip =
            "名前衝突時に連番追加で自動解消しない診断処理を実装するときに解除する。")]
    public void 名前衝突を自動的な連番追加では解消しません()
    {
        GeneratedCodeInspection
            .GenerateSources(ColumnNameCollisionExcelFilePath)
            .TypeDeclarations()
            .Should()
            .NotContain(it => it.Identifier.ValueText.Contains("CustomerId2"));
    }

    [Fact(
        Skip =
            "NameMappingsで生成名を変更して名前衝突を解消する処理を実装するときに解除する。")]
    public void NameMappingsで生成名を変更すると名前衝突を解消できます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                ColumnNameCollisionExcelFilePath,
                options => options.NameMappings["customer-id"] = "CustomerIdDash")
            .Should()
            .BeEmpty();
    }

    [Fact(
        Skip =
            "同じ生成型内で異なる種類のメンバー名が衝突した場合の診断を実装するときに解除する。")]
    public void 同じ生成型内の異なる種類のメンバー名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(BookMemberNameCollisionExcelFilePath)
            .Should()
            .Contain(it => it.IsError
                && it.GeneratedName == "SalesData"
                && it.SourceNames.Contains("sales_data")
                && it.SourceNames.Contains("sales-data"));
    }

    [Fact(
        Skip =
            "有効なCSharp識別子を生成できない名前のエラー診断を実装するときに解除する。")]
    public void 有効なCSharp識別子を生成できない名前を診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(InvalidNameExcelFilePath)
            .Should()
            .Contain(it => it.IsError
                && it.InvalidSourceName == "---");
    }
}
