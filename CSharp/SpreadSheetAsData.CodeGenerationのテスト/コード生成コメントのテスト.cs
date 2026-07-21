using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成コメントのテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

    [Fact]
    public void Book型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("BasicStructureBook")
            .SummaryText()
            .Should().Be("Excelブック「BasicStructure」を型付きで表します。");
    }

    [Fact]
    public void Sheet型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDataSheet")
            .SummaryText()
            .Should().Be("ワークシート「SalesData」を型付きで表します。");
    }

    [Fact]
    public void Table型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDetailTable")
            .SummaryText()
            .Should().Be("Excelテーブル「sales_detail」を型付きで表します。");
    }

    [Fact]
    public void 行データ型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDetail")
            .SummaryText()
            .Should().Be("Excelテーブル「sales_detail」の1行を表します。");
    }

    [Fact]
    public void Sheetプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("BasicStructureBook")
            .PropertyDeclaration("SalesData")
            .SummaryText()
            .Should().Be("ワークシート「SalesData」を取得します。");
    }

    [Fact]
    public void Tableプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("BasicStructureBook")
            .PropertyDeclaration("SalesDetail")
            .SummaryText()
            .Should().Be("Excelテーブル「sales_detail」を取得します。");
    }

    [Fact]
    public void SheetのTableプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDataSheet")
            .PropertyDeclaration("SalesDetail")
            .SummaryText()
            .Should().Be("Excelテーブル「sales_detail」を取得します。");
    }

    [Fact]
    public void 列プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDetail")
            .PropertyDeclaration("CustomerId")
            .SummaryText()
            .Should().Be("Excel列「customer_id」の値を取得または設定します。");
    }

    [Fact]
    public void ブックスコープの単一セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .TypeDeclaration("定義名Book")
            .PropertyDeclaration("MainCell")
            .SummaryText()
            .Should().Be("定義名「main_cell」が表すセルを取得します。");
    }

    [Fact]
    public void ブックスコープの複数セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .TypeDeclaration("定義名Book")
            .PropertyDeclaration("MainRange")
            .SummaryText()
            .Should().Be("定義名「main_range」が表すセル範囲を取得します。");
    }

    [Fact]
    public void シートローカルの単一セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .TypeDeclaration("SalesDataSheet")
            .PropertyDeclaration("LocalCell")
            .SummaryText()
            .Should().Be("ワークシート「sales_data」の定義名「local_cell」が表すセルを取得します。");
    }

    [Fact]
    public void シートローカルの複数セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .TypeDeclaration("SalesDataSheet")
            .PropertyDeclaration("LocalRange")
            .SummaryText()
            .Should().Be("ワークシート「sales_data」の定義名「local_range」が表すセル範囲を取得します。");
    }
}
