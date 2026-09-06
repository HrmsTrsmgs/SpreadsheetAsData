using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成名前設定のテスト
{
    const string SimpleNameMappingsExcelFilePath = @"TestData\コード生成\簡易名前置換.xlsx";
    const string ContextualNameMappingsExcelFilePath = @"TestData\コード生成\文脈付き名前置換.xlsx";
    const string DefinedNamesWithoutCollisionsExcelFilePath = @"TestData\コード生成\衝突なし\定義名.xlsx";

    [Fact]
    public void NameMappingsは自動名前変換より優先されます()
    {
        GeneratedCodeInspection
            .GenerateSources(
                SimpleNameMappingsExcelFilePath,
                options => options.NameMappings["cust_id"] = "CustomerID")
            .TypeDeclaration("SalesDetail")
            .PropertyDeclaration("CustomerID")
            .Should().NotBeNull();
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
            .TypeDeclaration("SalesDetail")
            .PropertyDeclaration("CustomerID")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは対象種類を指定せず同じ元名へ適用されます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            SimpleNameMappingsExcelFilePath,
            options => options.NameMappings["data"] = "MappedData");

        sources.TypeDeclaration("MappedDataSheet").Should().NotBeNull();
        sources.TypeDeclaration("簡易名前置換Book")
            .PropertyDeclaration("MappedData").Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは文脈付きキーで同じ元列名をテーブルごとに異なる名前へ変更できます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            ContextualNameMappingsExcelFilePath,
            options =>
                options.NameMappings = new()
                {
                    ["customers.id"] = "CustomerId",
                    ["products.id"] = "ProductId"
                });

        sources.TypeDeclaration("Customers")
            .PropertyDeclaration("CustomerId").Should().NotBeNull();
        sources.TypeDeclaration("Products")
            .PropertyDeclaration("ProductId").Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは文脈付きキーでブック定義名とシートローカル定義名を区別できます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
            ContextualNameMappingsExcelFilePath,
            options =>
                options.NameMappings = new()
                {
                    ["book.total"] = "GrandTotal",
                    ["sales_data.total"] = "SheetTotal"
                });

        sources.TypeDeclaration("文脈付き名前置換Book")
            .PropertyDeclaration("GrandTotal").Should().NotBeNull();
        sources.TypeDeclaration("SalesDataSheet")
            .PropertyDeclaration("SheetTotal").Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsは文脈付きキーを単純キーより優先します()
    {
        GeneratedCodeInspection
            .GenerateSources(
                ContextualNameMappingsExcelFilePath,
                options =>
                    options.NameMappings = new()
                    {
                        ["id"] = "MappedId",
                        ["customers.id"] = "CustomerId"
                    })
            .TypeDeclaration("Customers")
            .PropertyDeclaration("CustomerId")
            .Should().NotBeNull();
    }

    [Fact]
    public void NameMappingsで変更した生成Dataプロパティへ定義名の値を読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["book.main_cell"] = "PrimaryCell"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesWithoutCollisionsExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        object? tested = dataAccessor.PrimaryCell;

        tested.Should().Be("main");
    }

    [Fact]
    public void NameMappingsで変更した生成Dataプロパティから定義名へ書き込みます()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["book.main_cell"] = "PrimaryCell"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                temporaryFiles.Copy(DefinedNamesWithoutCollisionsExcelFilePath));
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        dataAccessor.PrimaryCell = "changed";

        bookAccessor.Replace(dataAccessor);

        (book.Cell["main_cell"].Value as object).Should().Be("changed");
    }
}
