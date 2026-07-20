using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class CSharp識別子生成のテスト
{
    const string CamelCaseIdentifierExcelFilePath = @"TestData\コード生成\salesReport.xlsx";
    const string JapaneseMixedNameExcelFilePath = @"TestData\コード生成\日本語混在名前.xlsx";

    [Theory]
    [InlineData("salesReport", "SalesReport")]
    [InlineData("salesData", "SalesData")]
    [InlineData("salesDetail", "SalesDetail")]
    [InlineData("customerId", "CustomerId")]
    public void camelCaseのExcel由来名はPascalCase識別子へ変換します(
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
    public void ASCII名は区切り文字を単語境界としてPascalCase識別子へ変換します(
        string excelName,
        string identifierBody)
    {
        WorkbookWrapperComponents
            .Identifier(excelName)
            .Should()
            .Be(identifierBody);
    }

    [Fact]
    public void ASCII名は全大文字の単語をPascalCase識別子へ正規化します()
    {
        WorkbookWrapperComponents
            .Identifier("SALES_DETAIL1")
            .Should()
            .Be("SalesDetail1");
    }

    [Fact]
    public void ASCII名は二文字頭字語を両方大文字の識別子へ変換します()
    {
        WorkbookWrapperComponents
            .Identifier("IO_stream")
            .Should()
            .Be("IOStream");
    }

    [Fact]
    public void ASCII名はIdを二文字頭字語の例外として変換します()
    {
        WorkbookWrapperComponents
            .Identifier("customer_ID")
            .Should()
            .Be("CustomerId");
    }

    [Theory]
    [InlineData("商品_明細", "商品_明細")]
    [InlineData("商品-明細", "商品_明細")]
    [InlineData("sales商品-detail", "Sales商品_detail")]
    public void 非ASCIIを含む名は内部の単語境界を推測しません(
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
            "非ASCIIを含む識別子の変換規則を列プロパティ名へ適用するときに解除する。")]
    public void 非ASCIIを含む列プロパティ名は識別子を使用します()
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

    [Theory]
    [InlineData('!')]
    [InlineData('"')]
    [InlineData('#')]
    [InlineData('%')]
    [InlineData('&')]
    [InlineData('\'')]
    [InlineData('*')]
    [InlineData(',')]
    [InlineData('.')]
    [InlineData('/')]
    [InlineData(':')]
    [InlineData(';')]
    [InlineData('?')]
    [InlineData('@')]
    [InlineData('\\')]
    public void OtherPunctuationに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('(')]
    [InlineData('[')]
    [InlineData('{')]
    public void OpenPunctuationに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData(')')]
    [InlineData(']')]
    [InlineData('}')]
    public void ClosePunctuationに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('+')]
    [InlineData('<')]
    [InlineData('=')]
    [InlineData('>')]
    [InlineData('|')]
    [InlineData('~')]
    public void MathSymbolに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('^')]
    [InlineData('`')]
    public void ModifierSymbolに分類されるASCII記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void CurrencySymbolに分類されるASCII記号はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price$rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('！')]
    [InlineData('＂')]
    [InlineData('＃')]
    [InlineData('％')]
    [InlineData('＆')]
    [InlineData('＇')]
    [InlineData('＊')]
    [InlineData('，')]
    [InlineData('．')]
    [InlineData('／')]
    [InlineData('：')]
    [InlineData('；')]
    [InlineData('？')]
    [InlineData('＠')]
    [InlineData('＼')]
    public void OtherPunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('（')]
    [InlineData('［')]
    [InlineData('｛')]
    public void OpenPunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('）')]
    [InlineData('］')]
    [InlineData('｝')]
    public void ClosePunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void DashPunctuationに分類される全角ASCII相当記号はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price－rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('＋')]
    [InlineData('＜')]
    [InlineData('＝')]
    [InlineData('＞')]
    [InlineData('｜')]
    [InlineData('～')]
    public void MathSymbolに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory]
    [InlineData('＾')]
    [InlineData('｀')]
    public void ModifierSymbolに分類される全角ASCII相当記号はアンダースコアへ置換します(
        char character)
    {
        WorkbookWrapperComponents
            .Identifier($"price{character}rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void CurrencySymbolに分類される全角ASCII相当記号はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price＄rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void OtherSymbolに分類される漢字構成記述文字はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("商品⿰明細")
            .Should()
            .Be("商品_明細");
    }

    [Fact]
    public void SpaceSeparatorに分類される全角空白はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price\u3000rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void Controlに分類される制御文字はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price\u0001rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void OtherNumberに分類される数値文字はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price\u00B2rate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void EnclosingMarkに分類される囲み結合記号はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price\u20DDrate")
            .Should()
            .Be("Price_rate");
    }

    [Fact]
    public void PrivateUseに分類される私用領域文字はアンダースコアへ置換します()
    {
        WorkbookWrapperComponents
            .Identifier("price\uE000rate")
            .Should()
            .Be("Price_rate");
    }

    [Theory(
        Skip =
            "数字から始まる名を有効なCSharp識別子へ補正するときに解除する。")]
    [InlineData("2026_sales", "_2026Sales")]
    [InlineData("2026商品", "_2026商品")]
    public void 数字から始まる名は有効なCSharp識別子へ補正します(
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
            "CSharpキーワードと同じExcel名をキーワードでない識別子へ変換するときに解除する。")]
    public void CSharpキーワードと同じExcel名はキーワードでない識別子へ変換します()
    {
        WorkbookWrapperComponents
            .Identifier("class")
            .Should()
            .Be("Class");
    }
}
