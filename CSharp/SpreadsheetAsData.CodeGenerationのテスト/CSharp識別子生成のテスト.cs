using FluentAssertions;
using Marimo.SpreadsheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadsheetAsData.CodeGeneration.Test;

public sealed class CSharp識別子生成のテスト
{
    const string キャメルケースIdentifierExcelFilePath = @"TestData\コード生成\salesReport.xlsx";
    const string JapaneseMixedNameExcelFilePath = @"TestData\コード生成\日本語混在名前.xlsx";

    [Theory]
    [InlineData("salesReport", "SalesReport")]
    [InlineData("salesData", "SalesData")]
    [InlineData("salesDetail", "SalesDetail")]
    [InlineData("customerId", "CustomerId")]
    public void キャメルケースのExcel由来名はPascalCase識別子へ変換します(
        string excelName,
        string identifierBody)
    {
        excelName.ToCSharpIdentifier()
            .Should().Be(identifierBody);
    }

    [Fact]
    public void キャメルケースブック名はPascalCaseのBook型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースIdentifierExcelFilePath))
            .TypeNames
            .Should().Contain("SalesReportBook");
    }

    [Fact]
    public void キャメルケースシート名はPascalCaseのSheet型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースIdentifierExcelFilePath))
            .TypeNames
            .Should().Contain("SalesDataSheet");
    }

    [Fact]
    public void キャメルケーステーブル名はPascalCaseのTable型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースIdentifierExcelFilePath))
            .TypeNames
            .Should().Contain("SalesDetailTable");
    }

    [Fact]
    public void キャメルケーステーブル名はPascalCaseの行データ型名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースIdentifierExcelFilePath))
            .TypeNames
            .Should().Contain("SalesDetail");
    }

    [Fact]
    public void キャメルケース列名はPascalCaseの行データプロパティ名へ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    キャメルケースIdentifierExcelFilePath))
            .GeneratedType("SalesDetail")
            .PropertyNames
            .Should().Contain("CustomerId");
    }

    [Theory]
    [InlineData("sales_detail", "SalesDetail")]
    [InlineData("sales-detail", "SalesDetail")]
    [InlineData("sales detail", "SalesDetail")]
    public void ASCII名は区切り文字を単語境界としてPascalCase識別子へ変換します(
        string excelName,
        string identifierBody)
    {
        excelName.ToCSharpIdentifier()
            .Should().Be(identifierBody);
    }

    [Fact]
    public void ASCII名は全大文字の単語をPascalCase識別子へ正規化します()
    {
        "SALES_DETAIL1".ToCSharpIdentifier()
            .Should().Be("SalesDetail1");
    }

    [Fact]
    public void ASCII名は二文字頭字語を両方大文字の識別子へ変換します()
    {
        "IO_stream".ToCSharpIdentifier()
            .Should().Be("IOStream");
    }

    [Fact]
    public void ASCII名はIdを二文字頭字語の例外として変換します()
    {
        "customer_ID".ToCSharpIdentifier()
            .Should().Be("CustomerId");
    }

    [Theory]
    [InlineData("商品_明細", "商品_明細")]
    [InlineData("商品-明細", "商品_明細")]
    [InlineData("sales商品-detail", "Sales商品_detail")]
    public void 非ASCIIを含む名は内部の単語境界を推測しません(
        string excelName,
        string identifierBody)
    {
        excelName.ToCSharpIdentifier()
            .Should().Be(identifierBody);
    }

    [Fact]
    public void 非ASCIIを含む列プロパティ名は識別子を使用します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    JapaneseMixedNameExcelFilePath))
            .GeneratedType("商品_明細")
            .PropertyNames
            .Should().Contain(["商品_id", "Sales商品_detail", "商品sales_detail"]);
    }

    [Theory]
    [InlineData("price/rate", "Price_rate")]
    [InlineData("price)rate", "Price_rate")]
    [InlineData("price$rate", "Price_rate")]
    [InlineData("price－rate", "Price_rate")]
    [InlineData("price＋rate", "Price_rate")]
    [InlineData("price\u3000rate", "Price_rate")]
    [InlineData("price\u0001rate", "Price_rate")]
    [InlineData("price\u00B2rate", "Price_rate")]
    [InlineData("商品⿰明細", "商品_明細")]
    public void 識別子に使えない文字はアンダースコアへ置換します(
        string excelName,
        string identifierBody)
    {
        excelName.ToCSharpIdentifier()
            .Should().Be(identifierBody);
    }

    [Theory]
    [InlineData("2026_sales", "_2026Sales")]
    [InlineData("2026商品", "_2026商品")]
    public void 数字から始まる名は有効なCSharp識別子へ補正します(
        string excelName,
        string identifierBody)
    {
        excelName.ToCSharpIdentifier()
            .Should().Be(identifierBody);
    }

    [Theory]
    [InlineData("\u203Fsales", "_\u203Fsales")]
    [InlineData("\u0301sales", "_\u0301sales")]
    [InlineData("\u0903sales", "_\u0903sales")]
    [InlineData("\u200Csales", "_\u200Csales")]
    public void 開始文字として使用できないカテゴリで始まる名は有効なCSharp識別子へ補正します(
        string excelName,
        string identifierBody)
    {
        excelName.ToCSharpIdentifier()
            .Should().Be(identifierBody);
    }

    [Fact]
    public void CSharpキーワードと同じExcel名はキーワードでない識別子へ変換します()
    {
        "class".ToCSharpIdentifier()
            .Should().Be("Class");
    }
}
