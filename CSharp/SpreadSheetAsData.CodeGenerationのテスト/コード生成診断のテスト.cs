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

    [Fact]
    public void 自動変換後に同じ列プロパティ名となる場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(ColumnNameCollisionExcelFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "CustomerId",
                    ["customer_id", "customer-id"]));
    }

    [Fact]
    public void 名前衝突を自動的な連番追加では解消しません()
    {
        GeneratedCodeInspection
            .GenerateSources(ColumnNameCollisionExcelFilePath)
            .TypeDeclarations()
            .Should().NotContain(it => it.Identifier.ValueText.Contains("CustomerId2"));
    }

    [Fact]
    public void NameMappingsで生成名を変更すると名前衝突を解消できます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                ColumnNameCollisionExcelFilePath,
                options => options.NameMappings["customer-id"] = "CustomerIdDash")
            .Should().BeEmpty();
    }

    [Fact]
    public void 同じ生成型内の異なる種類のメンバー名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(BookMemberNameCollisionExcelFilePath)
            .Should().Contain(it => it.IsError
                && it.GeneratedName == "SalesData"
                && it.SourceNames.Contains("sales_data")
                && it.SourceNames.Contains("sales-data"));
    }

    [Fact]
    public void 区切り文字だけのシート名はCSharp識別子を生成できないため診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(InvalidNameExcelFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "",
                    ["---"],
                    "---"));
    }
}
