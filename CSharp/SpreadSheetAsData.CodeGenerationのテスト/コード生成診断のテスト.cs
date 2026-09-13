using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成診断のテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
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
    public void 列プロパティ名が生成行データの型名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["sales_detail.customer_id"] = "SalesDetail" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "SalesDetail",
                    ["customer_id"]));
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
    public void 同じBook型のシートとテーブルの生成プロパティ名が衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["SalesData"] = "SalesDetail" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "SalesDetail",
                    ["SalesData", "sales_detail"]));
    }

    [Fact]
    public void シートの生成プロパティ名が生成BookのReadメソッド名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["SalesData"] = "Read" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Read",
                    ["SalesData"]));
    }

    [Fact]
    public void テーブルの生成プロパティ名が生成BookのReadメソッド名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["sales_detail"] = "Read" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Read",
                    ["sales_detail"]));
    }

    [Fact]
    public void シートの生成プロパティ名が生成BookのOpenメソッド名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["SalesData"] = "Open" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Open",
                    ["SalesData"]));
    }

    [Fact]
    public void テーブルの生成プロパティ名が生成BookのOpenメソッド名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["sales_detail"] = "Open" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Open",
                    ["sales_detail"]));
    }

    [Theory]
    [InlineData("Read")]
    [InlineData("Open")]
    [InlineData("Replace")]
    [InlineData("ValidateStructure")]
    public void ブックスコープの定義名が生成Bookのメソッド名と衝突した場合に診断します(string generatedName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["book.main_cell"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    generatedName,
                    ["main_cell"]));
    }

    [Fact]
    public void ブックスコープの定義名がWorkbookのCellプロパティ名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["book.main_cell"] = "Cell" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Cell",
                    ["main_cell"]));
    }

    [Fact]
    public void ブックスコープの定義名がWorkbookのRangeプロパティ名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["book.main_range"] = "Range" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Range",
                    ["main_range"]));
    }

    [Fact]
    public void ブックスコープの定義名がWorkbookのTablesプロパティ名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["book.main_cell"] = "Tables" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Tables",
                    ["main_cell"]));
    }

    [Fact]
    public void ブックスコープの定義名がWorkbookのDefinedNamesプロパティ名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["book.main_cell"] = "DefinedNames" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "DefinedNames",
                    ["main_cell"]));
    }

    [Fact]
    public void ブックスコープの定義名が生成Bookの型名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["book.main_cell"] = "定義名Book" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "定義名Book",
                    ["main_cell"]));
    }

    [Fact]
    public void シートローカルの定義名がWorksheetのCellプロパティ名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["sales_data.local_cell"] = "Cell" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Cell",
                    ["local_cell"]));
    }

    [Fact]
    public void シートローカルの定義名がWorksheetのRangeプロパティ名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["sales_data.local_range"] = "Range" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "Range",
                    ["local_range"]));
    }

    [Fact]
    public void シートローカルの定義名が生成Sheetの型名と衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\衝突なし\定義名.xlsx",
                options => options.NameMappings = new() { ["sales_data.local_cell"] = "SalesDataSheet" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "SalesDataSheet",
                    ["local_cell"]));
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
