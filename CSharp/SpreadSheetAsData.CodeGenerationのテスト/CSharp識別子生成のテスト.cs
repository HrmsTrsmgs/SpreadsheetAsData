using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class CSharp識別子生成のテスト
{
    const string AsciiNameConversionExcelFilePath = @"TestData\コード生成\ASCII名前変換.xlsx";
    const string CamelCaseIdentifierExcelFilePath = @"TestData\コード生成\salesReport.xlsx";
    const string JapaneseMixedNameExcelFilePath = @"TestData\コード生成\日本語混在名前.xlsx";

    [Theory]
    [InlineData("salesReport", "SalesReport")]
    [InlineData("salesData", "SalesData")]
    [InlineData("salesDetail", "SalesDetail")]
    [InlineData("customerId", "CustomerId")]
    public void camelCaseのExcel由来名はPascalCaseの識別子本文へ変換します(
        string excelName,
        string identifierBody)
    {
        WorkbookWrapperComponents
            .Identifier(excelName)
            .Should()
            .Be(identifierBody);
    }

    [Fact]
    public void camelCaseブック名はPascalCaseのBook型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    CamelCaseIdentifierExcelFilePath))
            .TypeNames
            .Should()
            .Contain("SalesReportBook");
    }

    [Fact]
    public void camelCaseシート名はPascalCaseのSheet型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    CamelCaseIdentifierExcelFilePath))
            .TypeNames
            .Should()
            .Contain("SalesDataSheet");
    }

    [Fact]
    public void camelCaseテーブル名はPascalCaseのTable型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    CamelCaseIdentifierExcelFilePath))
            .TypeNames
            .Should()
            .Contain("SalesDetailTable");
    }

    [Fact]
    public void camelCaseテーブル名はPascalCaseの行データ型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    CamelCaseIdentifierExcelFilePath))
            .TypeNames
            .Should()
            .Contain("SalesDetail");
    }

    [Fact]
    public void camelCase列名はPascalCaseの行データプロパティ名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    CamelCaseIdentifierExcelFilePath))
            .GeneratedType("SalesDetail")
            .PropertyNames
            .Should()
            .Contain("CustomerId");
    }

    [Theory]
    [InlineData("sales_detail", "SalesDetail")]
    [InlineData("sales-detail", "SalesDetail")]
    [InlineData("sales detail", "SalesDetail")]
    public void ASCII識別子本文は区切り文字を単語境界としてPascalCaseへ変換します(
        string excelName,
        string identifierBody)
    {
        WorkbookWrapperComponents
            .Identifier(excelName)
            .Should()
            .Be(identifierBody);
    }

    [Fact]
    public void ASCII識別子本文は全大文字の単語をPascalCaseへ正規化します()
    {
        WorkbookWrapperComponents
            .Identifier("SALES_DETAIL1")
            .Should()
            .Be("SalesDetail1");
    }

    [Fact]
    public void ASCII識別子本文は二文字頭字語を両方大文字で変換します()
    {
        WorkbookWrapperComponents
            .Identifier("IO_stream")
            .Should()
            .Be("IOStream");
    }

    [Fact]
    public void ASCII識別子本文はIdを二文字頭字語の例外として変換します()
    {
        WorkbookWrapperComponents
            .Identifier("customer_ID")
            .Should()
            .Be("CustomerId");
    }

    [Fact(
        Skip =
            "識別子本文の変換規則をTable型名と行データ型名の両方へ適用するときに解除する。")]
    public void Table型名と行データ型名は同じ識別子本文を使用します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .TypeNames
            .Should()
            .Contain(["SalesDetail", "SalesDetailTable"]);
    }

    [Fact(
        Skip =
            "識別子本文の変換規則を列プロパティ名へ適用するときに解除する。")]
    public void 列プロパティ名は識別子本文を使用します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .GeneratedType("SalesDetail")
            .PropertyNames
            .Should()
            .Contain("CustomerId");
    }

    [Theory(
        Skip =
            "非ASCIIを含む識別子本文の内部単語境界を推測しない変換規則を実装するときに解除する。")]
    [InlineData("商品_明細", "商品_明細")]
    [InlineData("商品-明細", "商品_明細")]
    [InlineData("sales商品-detail", "Sales商品_detail")]
    public void 非ASCIIを含む識別子本文は内部の単語境界を推測しません(
        string excelName,
        string identifierBody)
    {
        WorkbookWrapperComponents
            .Identifier(excelName)
            .Should()
            .Be(identifierBody);
    }

    [Fact(
        Skip =
            "非ASCIIを含む識別子本文の変換規則を列プロパティ名へ適用するときに解除する。")]
    public void 非ASCIIを含む列プロパティ名は識別子本文を使用します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    JapaneseMixedNameExcelFilePath))
            .GeneratedType("商品_明細")
            .PropertyNames
            .Should()
            .Contain(["商品_id", "Sales商品_detail", "商品sales_detail"]);
    }

    [Fact(
        Skip =
            "識別子本文に使用できない文字をアンダースコアへ置換するときに解除する。")]
    public void 識別子本文に使用できない文字はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("商品 明細")
            .Should()
            .Be("商品_明細");
    }

    [Theory(
        Skip =
            "数字から始まる識別子本文を有効なCSharp識別子へ補正するときに解除する。")]
    [InlineData("2026_sales", "_2026Sales")]
    [InlineData("2026商品", "_2026商品")]
    public void 数字から始まる識別子本文は有効なCSharp識別子へ補正します(
        string excelName,
        string identifierBody)
    {
        WorkbookWrapperComponents
            .Identifier(excelName)
            .Should()
            .Be(identifierBody);
    }

    [Fact(
        Skip =
            "CSharpキーワードと同じExcel名をキーワードでない識別子本文へ変換するときに解除する。")]
    public void CSharpキーワードと同じExcel名はキーワードでない識別子本文へ変換します()
    {
        WorkbookWrapperComponents
            .Identifier("class")
            .Should()
            .Be("Class");
    }
}
