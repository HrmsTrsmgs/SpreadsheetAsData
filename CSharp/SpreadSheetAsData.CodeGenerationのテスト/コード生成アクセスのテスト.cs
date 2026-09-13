using FluentAssertions;
using Marimo.SpreadSheetAsData;
using System.Reflection;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成アクセスのテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string WithoutTablesExcelFilePath = @"TestData\コード生成\テーブルなし.xlsx";
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\ブックスコープ\定義名.xlsx";
    const string DefinedNamesWithSheetScopeExcelFilePath = @"TestData\コード生成\シートローカル単一セル\定義名.xlsx";
    const string DefinedNamesWithoutCollisionsExcelFilePath = @"TestData\コード生成\衝突なし\定義名.xlsx";

    [Fact]
    public void Bookは各ワークシートを型付きプロパティとして公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var tested = generatedAssembly
            .GeneratedType("BasicStructureBook")
            .GetProperty("SalesData");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            generatedAssembly.GeneratedType("SalesDataSheet"));
    }

    [Fact]
    public void Bookはワークシートプロパティから型付きSheetを取得します()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        var tested = bookAccessor.SalesData as object;

        tested.Should().NotBeNull();
    }

    [Fact]
    public void Bookは各Excelテーブルを型付きプロパティとして公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var tested = generatedAssembly
            .GeneratedType("BasicStructureBook")
            .GetProperty("SalesDetail");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            generatedAssembly.GeneratedType("SalesDetailTable"));
    }

    [Fact]
    public void BookはExcelテーブルプロパティから型付きTableを取得します()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        var tested = bookAccessor.SalesDetail as object;

        tested.Should().NotBeNull();
    }

    [Fact]
    public void 生成されたBook型からWorkbookの非型付きAPIも使用できます()
    {
        using var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        tested.Sheets.Should().NotBeNull();
        tested.Tables.Should().NotBeNull();
        tested.Cell.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested["SalesData"].Should().NotBeNull();
    }

    [Fact]
    public void 生成されたBook型は必要なシートがないファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        // 生成元のSalesDataとProductMasterは、開くファイルには存在しません。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", DefinedNamesExcelFilePath);
        };

        // リフレクション経由の呼び出しでは、Openの例外がInnerExceptionに入ります。
        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は必要なテーブルがないファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        // SalesDataとProductMasterはありますが、sales_detailとProductListはありません。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", WithoutTablesExcelFilePath);
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は必要なシートがないStreamをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        // 生成元のSalesDataとProductMasterは、開くStreamには存在しません。
        using var stream = new MemoryStream(File.ReadAllBytes(DefinedNamesExcelFilePath));
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>("Open", stream);
        };

        // リフレクション経由の呼び出しでは、Openの例外がInnerExceptionに入ります。
        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型はStreamから開けます()
    {
        using var stream = new MemoryStream(
            File.ReadAllBytes(BasicStructureExcelFilePath));

        using var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook")
            .InvokeStaticMethod<Workbook>("Open", stream);

        tested.Sheets.Keys
            .Should().Equal("SalesData", "ProductMaster");
    }

    [Fact]
    public void 生成されたBookのReadとReplaceはCS0109をエラーとしてもコンパイルできます()
    {
        var action = () => GeneratedSourceCompiler.Compile(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath),
            warningsAsErrors: ["CS0109"]);

        action.Should().NotThrow();
    }

    [Fact]
    public void 生成されたBookはDataを型引数なしで読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        object? tested = dataAccessor.MainCell;

        tested.Should().Be("main");
    }

    [Fact]
    public void 生成されたBookはシートローカルの単一セル定義名をDataへ読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithSheetScopeExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesWithSheetScopeExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        object? tested = dataAccessor.LocalCell;

        tested.Should().Be(1d);
    }

    [Fact]
    public void 生成されたBookはシートローカルの複数セル定義名をDataへ読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesWithoutCollisionsExcelFilePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<IEnumerable<object?>> tested = dataAccessor.LocalRange;
        var rows = tested
            .Select(it => it.ToArray())
            .ToArray();

        rows[0].Should().Equal(1d, 10.5d);
        rows[1].Should().Equal(2d, 20.5d);
    }

    [Fact]
    public void 生成されたBookはData型を明記せず置換できます()
    {
        GeneratedSourceCompiler.Compile(
            [
                .. GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath),
                """
                namespace Generated;

                public static class Usage
                {
                    public static void Replace(定義名Book book) =>
                        book.Replace(new()
                        {
                            MainCell = "changed"
                        });
                }
                """
            ]);
    }

    [Fact]
    public void 生成されたBookはDataからシートローカルの単一セル定義名を置換します()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        var filePath = temporaryFiles.Copy(
            DefinedNamesWithSheetScopeExcelFilePath);

        using (var book = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesWithSheetScopeExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = book;
            dynamic dataAccessor = bookAccessor.Read();
            dataAccessor.LocalCell = 2d;

            bookAccessor.Replace(dataAccessor);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["sales_data"].Cell["local_cell"].Value as object)
            .Should().Be(2d);
    }

    [Fact]
    public void 生成されたBookはDataからシートローカルの複数セル定義名を置換します()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        var filePath = temporaryFiles.Copy(
            DefinedNamesWithoutCollisionsExcelFilePath);
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesWithoutCollisionsExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                filePath);
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<IEnumerable<object?>> replacement =
        [
            [3d, 30.5d],
            [4d, 40.5d]
        ];
        dataAccessor.LocalRange = replacement;

        bookAccessor.Replace(dataAccessor);

        book.Sheets["sales_data"].Range["local_range"].Values
            .Should().BeEquivalentTo(
                replacement,
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void Sheetはそのシートに属するExcelテーブルを型付きプロパティとして公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var sheetType = generatedAssembly.GeneratedType("SalesDataSheet");
        var tested = sheetType.GetProperty("SalesDetail");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            generatedAssembly.GeneratedType("SalesDetailTable"));
        sheetType.GetProperty("ProductList").Should().BeNull();
    }

    [Fact]
    public void SheetはExcelテーブルプロパティから型付きTableを取得します()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        var tested = bookAccessor.SalesData.SalesDetail as object;

        tested.Should().NotBeNull();
    }

    [Fact]
    public void 生成されたSheet型からWorksheetの非型付きAPIも使用できます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        Worksheet tested = bookAccessor.SalesData;

        tested.Name.Should().NotBeNull();
        tested.Book.Should().NotBeNull();
        tested.Cell.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested.Cells.Should().NotBeNull();
    }
}
