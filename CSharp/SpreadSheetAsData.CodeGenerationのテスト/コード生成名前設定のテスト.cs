using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成名前設定のテスト
{
    const string SimpleNameMappingsExcelFilePath = @"TestData\コード生成\簡易名前置換.xlsx";
    const string ContextNameMappingsExcelFilePath = @"TestData\コード生成\文脈付き名前置換.xlsx";

    [Fact(
        Skip =
            "NameMappingsを自動名前変換より優先して生成名へ適用するときに解除する。")]
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

    [Fact(
        Skip =
            "NameMappingsを要素種類や所属先を問わない元名の対応表として適用するときに解除する。")]
    public void NameMappingsは対象種類を指定せず同じ元名へ適用されます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            SimpleNameMappingsExcelFilePath,
            options => options.NameMappings["data"] = "MappedData");

        sources.TypeDeclaration("MappedDataSheet").Should().NotBeNull();
        sources.PropertyDeclaration("SimpleNameMappingsBook", "MappedData").Should().NotBeNull();
    }

    [Fact(
        Skip =
            "文脈付き名前設定で同じ元列名をテーブルごとに異なる名前へ変更するときに解除する。")]
    public void 文脈付き名前設定は同じ元列名をテーブルごとに異なる名前へ変更できます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            ContextNameMappingsExcelFilePath,
            options =>
            {
                options.ContextNameMappings["customers.id"] = "CustomerId";
                options.ContextNameMappings["products.id"] = "ProductId";
            });

        sources.PropertyDeclaration("CustomersRow", "CustomerId").Should().NotBeNull();
        sources.PropertyDeclaration("ProductsRow", "ProductId").Should().NotBeNull();
    }

    [Fact(
        Skip =
            "文脈付き名前設定でブック定義名とシートローカル定義名を区別するときに解除する。")]
    public void 文脈付き名前設定はブック定義名とシートローカル定義名を区別できます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            ContextNameMappingsExcelFilePath,
            options =>
            {
                options.ContextNameMappings["book.total"] = "GrandTotal";
                options.ContextNameMappings["sales_data.total"] = "SheetTotal";
            });

        sources.PropertyDeclaration("ContextNameMappingsBook", "GrandTotal").Should().NotBeNull();
        sources.PropertyDeclaration("SalesDataSheet", "SheetTotal").Should().NotBeNull();
    }

    [Fact(
        Skip =
            "文脈付き名前設定をNameMappingsより優先して適用するときに解除する。")]
    public void 文脈付き名前設定はNameMappingsより優先されます()
    {
        GeneratedCodeInspection
            .GenerateSources(
                ContextNameMappingsExcelFilePath,
                options =>
                {
                    options.NameMappings["id"] = "MappedId";
                    options.ContextNameMappings["customers.id"] = "CustomerId";
                })
            .PropertyDeclaration("CustomersRow", "CustomerId")
            .Should()
            .NotBeNull();
    }
}


