using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class CSharp識別子生成のテスト
{
    const string AsciiNameConversionExcelFilePath = @"TestData\コード生成\ASCII名前変換.xlsx";
    const string CamelCaseIdentifierExcelFilePath = @"TestData\コード生成\salesReport.xlsx";
    const string JapaneseMixedNameExcelFilePath = @"TestData\コード生成\日本語混在名前.xlsx";

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

    [Theory(
        Skip =
            "ASCIIシート名をPascalCaseのSheet型名へ自動変換する処理を実装するときに解除する。")]
    [InlineData("sales_detail", "SalesDetailSheet")]
    [InlineData("sales-detail", "SalesDetailSheet")]
    [InlineData("sales detail", "SalesDetailSheet")]
    [InlineData("SALES_DETAIL1", "SalesDetail1Sheet")]
    public void ASCIIシート名はPascalCaseのSheet型名へ変換します(
        string excelSheetName,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .TypeNames
            .Should()
            .Contain(
                generatedTypeName,
                "Excelシート名 {0} から生成される型名だから",
                excelSheetName);
    }

    [Theory(
        Skip =
            "ASCIIテーブル名をPascalCaseのTable型名へ自動変換する処理を実装するときに解除する。")]
    [InlineData("sales_detail", "SalesDetailTable")]
    [InlineData("sales_detail_dash", "SalesDetailDashTable")]
    public void ASCIIテーブル名はPascalCaseのTable型名へ変換します(
        string excelTableName,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .TypeNames
            .Should()
            .Contain(
                generatedTypeName,
                "Excelテーブル名 {0} から生成される型名だから",
                excelTableName);
    }

    [Theory(
        Skip =
            "ASCIIテーブル名をPascalCaseの行データ型名へ自動変換する処理を実装するときに解除する。")]
    [InlineData("sales_detail", "SalesDetail")]
    [InlineData("sales_detail_dash", "SalesDetailDash")]
    public void ASCIIテーブル名はPascalCaseの行データ型名へ変換します(
        string excelTableName,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .TypeNames
            .Should()
            .Contain(
                generatedTypeName,
                "Excelテーブル名 {0} から生成される型名だから",
                excelTableName);
    }

    [Theory(
        Skip =
            "ASCII列名をPascalCaseの行データプロパティ名へ自動変換する処理を実装するときに解除する。")]
    [InlineData("SalesDetail", "customer_id", "CustomerId")]
    [InlineData("SalesDetail", "url_value", "UrlValue")]
    [InlineData("SalesDetailDash", "xml_id", "XmlId")]
    [InlineData("SalesDetailDash", "api_url", "ApiUrl")]
    public void ASCII列名はPascalCaseの行データプロパティ名へ変換します(
        string dataTypeName,
        string excelColumnName,
        string generatedPropertyName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .GeneratedType(dataTypeName)
            .PropertyNames
            .Should()
            .Contain(
                generatedPropertyName,
                "Excel列名 {0} から生成されるプロパティ名だから",
                excelColumnName);
    }

    [Theory(
        Skip =
            "非ASCIIを含むシート名の内部単語境界を推測しない識別子変換を実装するときに解除する。")]
    [InlineData("商品_明細", "商品_明細Sheet")]
    [InlineData("商品-明細", "商品_明細Sheet")]
    [InlineData("sales商品-detail", "Sales商品_detailSheet")]
    public void 日本語混在シート名は内部の単語境界を推測しません(
        string excelSheetName,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    JapaneseMixedNameExcelFilePath))
            .TypeNames
            .Should()
            .Contain(
                generatedTypeName,
                "Excelシート名 {0} から生成される型名だから",
                excelSheetName);
    }

    [Theory(
        Skip =
            "非ASCIIを含む列名の内部単語境界を推測しない識別子変換を実装するときに解除する。")]
    [InlineData("商品_id", "商品_id")]
    [InlineData("sales商品_detail", "Sales商品_detail")]
    [InlineData("商品sales_detail", "商品sales_detail")]
    public void 日本語混在列名は内部の単語境界を推測しません(
        string excelColumnName,
        string generatedPropertyName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    JapaneseMixedNameExcelFilePath))
            .GeneratedType("商品_明細")
            .PropertyNames
            .Should()
            .Contain(
                generatedPropertyName,
                "Excel列名 {0} から生成されるプロパティ名だから",
                excelColumnName);
    }

    [Fact(
        Skip =
            "非ASCIIを含む名前の使用できない識別子文字をアンダースコアへ置換するときに解除する。")]
    public void 日本語混在列名の使用できない識別子文字はアンダースコアへ置換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    JapaneseMixedNameExcelFilePath))
            .GeneratedType("商品_明細")
            .PropertyNames
            .Should()
            .Contain(
                "商品_明細",
                "Excel列名 商品 明細 から生成されるプロパティ名だから");
    }

    [Theory(
        Skip =
            "数字から始まるシート名を有効なCSharp識別子へ補正するときに解除する。")]
    [InlineData("2026_sales", "_2026SalesSheet")]
    [InlineData("2026商品", "_2026商品Sheet")]
    public void 数字から始まるシート名は有効なCSharp識別子へ補正します(
        string excelSheetName,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .TypeNames
            .Should()
            .Contain(
                generatedTypeName,
                "Excelシート名 {0} から生成される型名だから",
                excelSheetName);
    }

    [Fact(
        Skip =
            "自動変換名がCSharpキーワードにならないようPascalCaseへ変換するときに解除する。")]
    public void 自動変換ではキーワードもPascalCaseへ変換します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    AsciiNameConversionExcelFilePath))
            .TypeNames
            .Should()
            .Contain("ClassSheet");
    }
}
