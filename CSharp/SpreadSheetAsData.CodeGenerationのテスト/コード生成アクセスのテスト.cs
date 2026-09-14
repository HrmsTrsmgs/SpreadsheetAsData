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
    const string IntegratedExcelFilePath = @"TestData\コード生成\統合.xlsx";
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
    public void 生成されたBook型は必要なテーブル列がないファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        // シートとテーブルは同じですが、sales_detailのDescription列がありません。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\列不足.xlsx");
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は必要なブックスコープの定義名がないファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("定義名Book");

        // シートとテーブル、main_cellはありますが、main_rangeとtotalはありません。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", IntegratedExcelFilePath);
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は単一セルのブック定義名が複数セルに変わったファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("定義名Book");

        // main_cellの参照先だけがE1からE1:F1へ広がり、他の構造と値は同じです。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\単一セル定義名の範囲化\定義名.xlsx");
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は必要なシートローカルの定義名がないファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesWithSheetScopeExcelFilePath))
            .GeneratedType("定義名Book");

        // シート、テーブル、ブックスコープの定義名は同じですが、sales_dataのlocal_cellはありません。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", DefinedNamesExcelFilePath);
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は単一セルのシートローカル定義名が複数セルに変わったファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesWithSheetScopeExcelFilePath))
            .GeneratedType("定義名Book");

        // sales_dataのlocal_cellだけがA2からA2:B2へ広がり、他の構造と値は同じです。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\シートローカル単一セル定義名の範囲化\定義名.xlsx");
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は同名のブック定義名があっても必要なシートローカル定義名がなければOpenに失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    @"TestData\コード生成\定義名.xlsx",
                    options => options.NameMappings = new() { ["sales_data.total"] = "SheetTotal" }))
            .GeneratedType("定義名Book");

        // totalはブックスコープにだけ存在し、sales_dataではsheet_totalという別名です。
        // その他のシート、テーブル、定義名は生成元と一致します。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", DefinedNamesWithoutCollisionsExcelFilePath);
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は別シートの同名定義名を必要なシートローカル定義名の代わりにしません()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesWithSheetScopeExcelFilePath))
            .GeneratedType("定義名Book");

        // local_cellの参照先は同じsales_data!A2ですが、所属はproduct_masterです。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\別シートローカル定義名\定義名.xlsx");
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Theory]
    [InlineData(WithoutTablesExcelFilePath, "テーブルなしBook", BasicStructureExcelFilePath)]
    [InlineData(@"TestData\コード生成\列不足.xlsx", "列不足Book", BasicStructureExcelFilePath)]
    [InlineData(DefinedNamesExcelFilePath, "定義名Book", DefinedNamesWithSheetScopeExcelFilePath)]
    public void 生成されたBook型は生成元になかったテーブルや列や定義名が追加されてもOpenできます(
        string sourceExcelFilePath,
        string bookTypeName,
        string openedExcelFilePath)
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(sourceExcelFilePath))
            .GeneratedType(bookTypeName);

        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>("Open", openedExcelFilePath);
        };

        tested.Should().NotThrow();
    }

    [Theory]
    [InlineData(@"TestData\コード生成\単一セル定義名の範囲化\定義名.xlsx", DefinedNamesExcelFilePath)]
    [InlineData(@"TestData\コード生成\シートローカル単一セル定義名の範囲化\定義名.xlsx", DefinedNamesWithSheetScopeExcelFilePath)]
    public void 生成されたBook型は複数セルの定義名が単一セルに変わってもOpenできます(
        string sourceExcelFilePath,
        string openedExcelFilePath)
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(sourceExcelFilePath))
            .GeneratedType("定義名Book");

        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>("Open", openedExcelFilePath);
        };

        tested.Should().NotThrow();
    }

    [Fact]
    public void 生成されたBook型は定義名の参照位置だけが変わってもOpenできます()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(DefinedNamesWithoutCollisionsExcelFilePath))
            .GeneratedType("定義名Book");

        // main_cellはE1からF1へ、local_cellはA2からB2へ移動しています。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\定義名参照位置変更\定義名.xlsx");
        };

        tested.Should().NotThrow();
    }

    [Theory]
    [InlineData(@"TestData\コード生成\データ行減少.xlsx")]
    [InlineData(@"TestData\コード生成\データ行なし.xlsx")]
    public void 生成されたBook型はテーブルのデータ行が減っても空になってもOpenできます(string excelFilePath)
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>("Open", excelFilePath);
        };

        tested.Should().NotThrow();
    }

    [Fact]
    public void 生成されたBook型は必要なテーブルが別シートへ移動したファイルをOpenすると失敗します()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        // シート名とテーブル名は揃っていますが、二つのテーブルの所属シートが逆です。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\テーブル所属シート変更.xlsx");
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型は別テーブルの同名列を必要な列の代わりにしません()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        // DescriptionはProductListにだけあり、sales_detailにはありません。
        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\別テーブルに同名列.xlsx");
        };

        tested.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();
    }

    [Fact]
    public void 生成されたBook型はテーブルの列順だけが変わってもOpenできます()
    {
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook");

        var tested = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>(
                "Open", @"TestData\コード生成\列順変更.xlsx");
        };

        tested.Should().NotThrow();
    }

    [Fact]
    public void 生成されたBook型は構造の検証でOpenに失敗するとファイルを解放します()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        var excelFilePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesWithSheetScopeExcelFilePath))
            .GeneratedType("定義名Book");

        var openBook = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>("Open", excelFilePath);
        };

        openBook.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();

        var tested = () =>
        {
            using var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        };

        tested.Should().NotThrow();
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
    public void 生成されたBook型は構造の検証でOpenに失敗しても呼び出し側のStreamを閉じません()
    {
        using var stream = new MemoryStream(File.ReadAllBytes(DefinedNamesExcelFilePath));
        var generatedType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesWithSheetScopeExcelFilePath))
            .GeneratedType("定義名Book");

        var openBook = () =>
        {
            using var book = generatedType.InvokeStaticMethod<Workbook>("Open", stream);
        };

        openBook.Should().Throw<TargetInvocationException>()
            .WithInnerException<InvalidDataException>();

        stream.CanRead.Should().BeTrue();
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
