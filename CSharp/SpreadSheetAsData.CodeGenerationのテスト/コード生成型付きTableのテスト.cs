using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型付きTableのテスト : IDisposable
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";

    readonly TemporaryExcelFiles temporaryFiles = new();

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void 生成されたTableはPOCOをExcel上の順序で列挙します()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        IEnumerable<object> salesDetail = bookAccessor.SalesDetail;
        var tested = salesDetail.ToArray();

        tested.Select(it => PropertyValue(it, "CustomerId"))
            .Should().Equal(1, 2);
        tested.Select(it => PropertyValue(it, "Amount"))
            .Should().Equal(10.5, 20.5);
        tested.Select(it => PropertyValue(it, "Description"))
            .Should().Equal("a", "b");
    }

    [Fact]
    public void 生成されたTableはnullableプロパティへ値と空白を読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedInstance<Workbook>("テーブルBook");

        dynamic bookAccessor = book;
        IEnumerable<object> rows = bookAccessor.空白数値マッピング;
        var tested = rows.ToArray();

        tested.Select(it => PropertyValue(it, "数値"))
            .Should().Equal(0, null);
        tested.Select(it => PropertyValue(it, "小数"))
            .Should().Equal(1.5, null);
        tested.Select(it => PropertyValue(it, "真偽値"))
            .Should().Equal(false, null);
    }

    [Fact]
    public void 生成されたTableは空白セルをstringプロパティの空文字列として読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedInstance<Workbook>(
                "テーブルBook", temporaryFiles.Copy(@"TestData\テーブル.xlsx"));
        book.Tables["型付き行マッピング"].Rows.First()["文字列"].Value = null;

        dynamic bookAccessor = book;
        IEnumerable<object> rows = bookAccessor.型付き行マッピング;

        rows.Select(it => PropertyValue(it, "文字列"))
            .Should().Equal("", "たちつてと", "なにぬねの");
    }

    [Fact]
    public void 生成されたTableをTableとして扱うと非型付き行を利用できます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        Table tested = bookAccessor.SalesDetail;

        tested.Rows.Should().NotBeEmpty();
    }

    [Fact]
    public void 生成されたTable型からTableの構造情報を使用できます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        dynamic bookAccessor = book;
        Table tested = bookAccessor.SalesDetail;

        tested.Name.Should().Be("sales_detail");
        tested.Worksheet.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested.Columns.Should().NotBeEmpty();
    }

    [Fact]
    public void 生成された行データ型は利用者定義POCOと同じ変換規則で読み込まれます()
    {
        (object? CustomerId, object? Amount, object? Description)[] generatedRows;

        using (var generatedWorkbook = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           BasicStructureExcelFilePath))
                   .GeneratedInstance<Workbook>("BasicStructureBook"))
        {
            dynamic bookAccessor = generatedWorkbook;
            IEnumerable<object> salesDetail = bookAccessor.SalesDetail;
            generatedRows = salesDetail.Select(ReadGeneratedRow).ToArray();
        }

        using var book = Workbook.Open(BasicStructureExcelFilePath);

        generatedRows.Should().Equal(
                book.ReadTable<ReadTableComparison>("sales_detail")
                    .Select(ReadHandWrittenRow));
    }

    [Fact]
    public void 生成されたDataはExcelテーブルの行データを読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<object> tested = dataAccessor.SalesDetail;

        tested.Select(it => PropertyValue(it, "CustomerId"))
            .Should().Equal(1, 2);
    }

    [Fact]
    public void 生成されたDataはテーブルのnullableプロパティへ値と空白を読み込みます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedInstance<Workbook>("テーブルBook");

        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<object> rows = dataAccessor.空白数値マッピング;
        var tested = rows.ToArray();

        tested.Select(it => PropertyValue(it, "数値"))
            .Should().Equal(0, null);
        tested.Select(it => PropertyValue(it, "小数"))
            .Should().Equal(1.5, null);
        tested.Select(it => PropertyValue(it, "真偽値"))
            .Should().Equal(false, null);
    }

    [Fact]
    public void 生成されたDataからExcelテーブルの行データを置換します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));
        var rowType = generatedAssembly.GeneratedType("SalesDetail");
        using var tested = generatedAssembly.GeneratedInstance<Workbook>(
            "BasicStructureBook",
            temporaryFiles.Copy(BasicStructureExcelFilePath));
        dynamic bookAccessor = tested;
        dynamic dataAccessor = bookAccessor.Read();
        dynamic replacement = CreateRows(
            rowType,
            (10, 1.5, "first"),
            (20, 2.5, "second"));
        dataAccessor.SalesDetail = replacement;

        bookAccessor.Replace(dataAccessor);

        tested.Tables["sales_detail"].Rows
            .Select(it => it["customer_id"].Value)
            .Should().Equal(10d, 20d);
    }

    [Fact]
    public void 生成されたTable型のReplaceで型付き行を書き込めます()
    {
        var filePath = temporaryFiles.Copy(BasicStructureExcelFilePath);
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));
        var rowType = generatedAssembly.GeneratedType("SalesDetail");

        using (var book = generatedAssembly.GeneratedInstance<Workbook>(
                   "BasicStructureBook",
                   filePath))
        {
            dynamic bookAccessor = book;
            dynamic generatedRows = CreateRows(
                rowType,
                (10, 1.5, "first"),
                (20, 2.5, "second"));

            bookAccessor.SalesDetail.Replace(generatedRows);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.ReadTable<ReadTableComparison>("sales_detail")
            .Should().BeEquivalentTo(
                [
                    new ReadTableComparison
                    {
                        CustomerId = 10,
                        Amount = 1.5,
                        Description = "first"
                    },
                    new ReadTableComparison
                    {
                        CustomerId = 20,
                        Amount = 2.5,
                        Description = "second"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void 生成されたTableのReplaceはnullableプロパティの値とnullを書き込みます()
    {
        const string excelFilePath = @"TestData\テーブル.xlsx";
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedInstance<Workbook>("テーブルBook", temporaryFiles.Copy(excelFilePath));

        dynamic bookAccessor = book;
        dynamic replacement = Enumerable.ToArray(bookAccessor.空白数値マッピング);
        replacement[0].数値 = null;
        replacement[0].小数 = null;
        replacement[0].真偽値 = null;
        replacement[1].数値 = 10;
        replacement[1].小数 = 2.5;
        replacement[1].真偽値 = true;

        bookAccessor.空白数値マッピング.Replace(replacement);

        var tested = book.Tables["空白数値マッピング"].Rows.ToArray();
        tested.Select(it => it["数値"].Value as object)
            .Should().Equal(new BlankValue(), 10d);
        tested.Select(it => it["小数"].Value as object)
            .Should().Equal(new BlankValue(), 2.5);
        tested.Select(it => it["真偽値"].Value as object)
            .Should().Equal(new BlankValue(), true);
    }

    [Fact]
    public void 生成されたDataからテーブルのnullableプロパティの値とnullを書き込みます()
    {
        const string excelFilePath = @"TestData\テーブル.xlsx";
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedInstance<Workbook>("テーブルBook", temporaryFiles.Copy(excelFilePath));

        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        dynamic replacement = Enumerable.ToArray(dataAccessor.空白数値マッピング);
        replacement[0].数値 = null;
        replacement[0].小数 = null;
        replacement[0].真偽値 = null;
        replacement[1].数値 = 10;
        replacement[1].小数 = 2.5;
        replacement[1].真偽値 = true;
        dataAccessor.空白数値マッピング = replacement;

        bookAccessor.Replace(dataAccessor);

        var tested = book.Tables["空白数値マッピング"].Rows.ToArray();
        tested.Select(it => it["数値"].Value as object)
            .Should().Equal(new BlankValue(), 10d);
        tested.Select(it => it["小数"].Value as object)
            .Should().Equal(new BlankValue(), 2.5);
        tested.Select(it => it["真偽値"].Value as object)
            .Should().Equal(new BlankValue(), true);
    }

    static (object? CustomerId, object? Amount, object? Description) ReadGeneratedRow(object row) =>
        (
            PropertyValue(row, "CustomerId"),
            PropertyValue(row, "Amount"),
            PropertyValue(row, "Description")
        );

    static object? PropertyValue(object source, string propertyName) =>
        source.GetType()
            .GetProperty(propertyName)
            ?.GetValue(source);

    static (object? CustomerId, object? Amount, object? Description) ReadHandWrittenRow(ReadTableComparison row) =>
        (
            row.CustomerId,
            row.Amount,
            row.Description
        );

    static Array CreateRows(
        Type rowType,
        params (int CustomerId, double Amount, string Description)[] sourceRows)
    {
        var rows = Array.CreateInstance(rowType, sourceRows.Length);

        for (var index = 0; index < sourceRows.Length; index++)
        {
            rows.SetValue(CreateRow(rowType, sourceRows[index]), index);
        }

        return rows;
    }

    static object CreateRow(
        Type rowType,
        (int CustomerId, double Amount, string Description) source)
    {
        var row = Activator.CreateInstance(rowType)
            ?? throw new InvalidOperationException(rowType.Name);

        rowType.GetProperty("CustomerId")?.SetValue(row, source.CustomerId);
        rowType.GetProperty("Amount")?.SetValue(row, source.Amount);
        rowType.GetProperty("Description")?.SetValue(row, source.Description);

        return row;
    }

    sealed class ReadTableComparison
    {
        [SpreadSheetName("customer_id")]
        public int CustomerId { get; set; }

        public double Amount { get; set; }

        public string Description { get; set; } = "";
    }
}
