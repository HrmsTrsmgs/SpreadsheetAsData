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
    public void 書式文字を含む定義名から生成Dataプロパティへ値を読み込みます()
    {
        const string excelFilePath = @"TestData\コード生成\FormattingCharacter.xlsx";
        using var book = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedInstance<Workbook>("FormattingCharacterBook", excelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        object tested = dataAccessor.CustomerName;

        tested.Should().Be("直接文字列");
    }

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
    public void NameMappingsはExcelテーブル名へ適用されます()
    {
        var tested = GeneratedCodeInspection.SyntaxFrom(
            GeneratedCodeInspection.GenerateSources(
                SimpleNameMappingsExcelFilePath,
                options => options.NameMappings["sales_detail"] = "OrderLine"));

        tested.TypeNames.Should().Contain(["OrderLine", "OrderLineTable"]);
        tested.GeneratedType("簡易名前置換Book")
            .PropertyNames.Should().Contain("OrderLine");
        tested.GeneratedType("DataSheet")
            .PropertyNames.Should().Contain("OrderLine");
        tested.GeneratedType("簡易名前置換Data")
            .PropertyNames.Should().Contain("OrderLine");
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
    public void NameMappingsで変更した生成DataプロパティへExcelテーブルの行データを読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SimpleNameMappingsExcelFilePath,
                    options =>
                        options.NameMappings = new()
                        {
                            ["sales_detail"] = "OrderLine",
                            ["sales_detail.sales_detail"] = "SalesDetailValue"
                        }))
            .GeneratedInstance<Workbook>(
                "簡易名前置換Book",
                SimpleNameMappingsExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<object> rows = dataAccessor.OrderLine;
        dynamic rowAccessor = rows.Single();
        object? tested = rowAccessor.CustId;

        tested.Should().Be(1);
    }

    [Fact]
    public void NameMappingsで変更した生成DataプロパティからExcelテーブルへ書き込みます()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    SimpleNameMappingsExcelFilePath,
                    options =>
                        options.NameMappings = new()
                        {
                            ["sales_detail"] = "OrderLine",
                            ["sales_detail.sales_detail"] = "SalesDetailValue"
                        }))
            .GeneratedInstance<Workbook>(
                "簡易名前置換Book",
                temporaryFiles.Copy(SimpleNameMappingsExcelFilePath));
        dynamic bookAccessor = book;
        var data = bookAccessor.Read();
        data.OrderLine = Enumerable.ToArray(data.OrderLine);
        data.OrderLine[0].CustId = 2;

        bookAccessor.Replace(data);

        (book.Tables["sales_detail"].Rows.Single()["cust_id"].Value as object)
            .Should().Be(2d);
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

    [Fact]
    public void NameMappingsで変更した生成Data範囲プロパティへ定義名の値を読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["book.main_range"] = "PrimaryRange"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesWithoutCollisionsExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<IEnumerable<object?>> tested = dataAccessor.PrimaryRange;
        var rows = tested
            .Select(it => it.ToArray())
            .ToArray();

        rows[0].Should().Equal("main", "range");
        rows[1].Should().Equal(100d, 200d);
    }

    [Fact]
    public void NameMappingsで変更した生成Data範囲プロパティから定義名へ書き込みます()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["book.main_range"] = "PrimaryRange"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                temporaryFiles.Copy(DefinedNamesWithoutCollisionsExcelFilePath));
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<IEnumerable<object?>> replacement =
        [
            ["changed", "values"],
            [300d, 400d]
        ];
        dataAccessor.PrimaryRange = replacement;

        bookAccessor.Replace(dataAccessor);

        book.Range["main_range"].Values
            .Should().BeEquivalentTo(
                replacement,
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void 文脈付きNameMappingsで変更した生成Dataプロパティへシートローカル定義名の値を読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["sales_data.local_cell"] = "PrimaryLocalCell"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesWithoutCollisionsExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        object? tested = dataAccessor.PrimaryLocalCell;

        tested.Should().Be(1d);
    }

    [Fact]
    public void 文脈付きNameMappingsで変更した生成Dataプロパティからシートローカル定義名へ書き込みます()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["sales_data.local_cell"] = "PrimaryLocalCell"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                temporaryFiles.Copy(DefinedNamesWithoutCollisionsExcelFilePath));
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        dataAccessor.PrimaryLocalCell = 2d;

        bookAccessor.Replace(dataAccessor);

        (book.Sheets["sales_data"].Cell["local_cell"].Value as object)
            .Should().Be(2d);
    }

    [Fact]
    public void 文脈付きNameMappingsで変更した生成Data範囲プロパティへシートローカル定義名の値を読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["sales_data.local_range"] = "PrimaryLocalRange"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesWithoutCollisionsExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<IEnumerable<object?>> tested = dataAccessor.PrimaryLocalRange;
        var rows = tested
            .Select(it => it.ToArray())
            .ToArray();

        rows[0].Should().Equal(1d, 10.5d);
        rows[1].Should().Equal(2d, 20.5d);
    }

    [Fact]
    public void 文脈付きNameMappingsで変更した生成Data範囲プロパティからシートローカル定義名へ書き込みます()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath,
                    options => options.NameMappings["sales_data.local_range"] = "PrimaryLocalRange"))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                temporaryFiles.Copy(DefinedNamesWithoutCollisionsExcelFilePath));
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<IEnumerable<object?>> replacement =
        [
            [3d, 30.5d],
            [4d, 40.5d]
        ];
        dataAccessor.PrimaryLocalRange = replacement;

        bookAccessor.Replace(dataAccessor);

        book.Sheets["sales_data"].Range["local_range"].Values
            .Should().BeEquivalentTo(
                replacement,
                options => options.WithStrictOrdering());
    }
}
