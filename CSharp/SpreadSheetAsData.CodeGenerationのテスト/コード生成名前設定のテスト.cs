using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成名前設定のテスト
{
    const string SimpleNameMappingsExcelFilePath = @"TestData\コード生成\簡易名前置換.xlsx";
    const string ContextualNameMappingsExcelFilePath = @"TestData\コード生成\文脈付き名前置換.xlsx";

    [Fact]
    public void NameMappingsは自動名前変換より優先されます()
    {
        GeneratedCodeInspection
            .GenerateSources(
                SimpleNameMappingsExcelFilePath,
                options => options.NameMappings["cust_id"] = "CustomerID")
            .PropertyDeclaration("SalesDetail", "CustomerID")
            .Should()
            .NotBeNull();
    }

    [Fact]
    public void NameMappingsは辞書を代入して設定できます()
    {
        GeneratedCodeInspection
            .GenerateSources(
                SimpleNameMappingsExcelFilePath,
                options =>
                    options.NameMappings = new()
                    {
                        ["cust_id"] = "CustomerID"
                    })
            .PropertyDeclaration("SalesDetail", "CustomerID")
            .Should()
            .NotBeNull();
    }

    [Fact]
    public void NameMappingsは対象種類を指定せず同じ元名へ適用されます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            SimpleNameMappingsExcelFilePath,
            options => options.NameMappings["data"] = "MappedData");

        sources.TypeDeclaration("MappedDataSheet").Should().NotBeNull();
        sources.PropertyDeclaration("簡易名前置換Book", "MappedData").Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは文脈付きキーで同じ元列名をテーブルごとに異なる名前へ変更できます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            ContextualNameMappingsExcelFilePath,
            options =>
            {
                options.NameMappings["customers.id"] = "CustomerId";
                options.NameMappings["products.id"] = "ProductId";
            });

        sources.PropertyDeclaration("Customers", "CustomerId").Should().NotBeNull();
        sources.PropertyDeclaration("Products", "ProductId").Should().NotBeNull();
    }

    [Fact(
        Skip =
            "NameMappingsの文脈付きキーでブック定義名とシートローカル定義名を区別するときに解除する。")]
    public void NameMappingsは文脈付きキーでブック定義名とシートローカル定義名を区別できます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            ContextualNameMappingsExcelFilePath,
            options =>
            {
                options.NameMappings["book.total"] = "GrandTotal";
                options.NameMappings["sales_data.total"] = "SheetTotal";
            });

        sources.PropertyDeclaration("文脈付き名前置換Book", "GrandTotal").Should().NotBeNull();
        sources.PropertyDeclaration("SalesDataSheet", "SheetTotal").Should().NotBeNull();
    }

    [Fact(
        Skip =
            "NameMappingsの文脈付きキーを単純キーより優先して適用するときに解除する。")]
    public void NameMappingsは文脈付きキーを単純キーより優先します()
    {
        GeneratedCodeInspection
            .GenerateSources(
                ContextualNameMappingsExcelFilePath,
                options =>
                {
                    options.NameMappings["id"] = "MappedId";
                    options.NameMappings["customers.id"] = "CustomerId";
                })
            .PropertyDeclaration("Customers", "CustomerId")
            .Should()
            .NotBeNull();
    }
}


