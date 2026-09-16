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
    public void 生成されたTableは混在列の値を元の型のまま読み込みます()
    {
        var excelFilePath = temporaryFiles.Copy(BasicStructureExcelFilePath);
        using (var book = Workbook.Open(excelFilePath))
        {
            book.Tables["sales_detail"].Rows.First()["Description"].Value = 1d;
            book.Save();
        }

        using var sourceBook = Workbook.Open(excelFilePath);
        var tested = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedInstance<IEnumerable<object>>("SalesDetailTable", sourceBook.Tables["sales_detail"]);

        tested.Select(it => PropertyValue(it, "Description")).Should().Equal(1d, "b");
    }

    [Fact]
    public void 生成されたTableは改行を含む列名からも値を読み込みます()
    {
        const string excelFilePath = @"TestData\コード生成\改行列名.xlsx";
        using var sourceBook = Workbook.Open(excelFilePath, validate: true);
        var tested = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedInstance<IEnumerable<object>>("SalesDetailTable", sourceBook.Tables["sales_detail"]);

        tested.Select(it => PropertyValue(it, "Customer_id")).Should().Equal(1, 2);
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
    public void 生成されたTableはReadTableで取得した型付きTableと同じ行データを返します()
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
    public void 生成されたBookのReadは生成されたTableと同じ行データを返します()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");
        dynamic bookAccessor = book;
        dynamic dataAccessor = bookAccessor.Read();
        IEnumerable<object> tested = dataAccessor.SalesDetail;
        IEnumerable<object> tableRows = bookAccessor.SalesDetail;

        tested.Select(ReadGeneratedRow)
            .Should().Equal(tableRows.Select(ReadGeneratedRow));
    }

    [Fact]
    public void 生成されたBookのReplaceは生成されたTableのReplaceと同じセル値を書き込みます()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));
        var rowType = generatedAssembly.GeneratedType("SalesDetail");
        using var tested = generatedAssembly.GeneratedInstance<Workbook>(
            "BasicStructureBook",
            temporaryFiles.Copy(BasicStructureExcelFilePath));
        using var tableBook = generatedAssembly.GeneratedInstance<Workbook>(
            "BasicStructureBook",
            temporaryFiles.Copy(BasicStructureExcelFilePath));
        dynamic bookAccessor = tested;
        dynamic tableBookAccessor = tableBook;
        dynamic dataAccessor = bookAccessor.Read();
        dynamic replacement = CreateRows(
            rowType,
            (10, 1.5, "first"),
            (20, 2.5, "second"));
        dataAccessor.SalesDetail = replacement;

        bookAccessor.Replace(dataAccessor);
        tableBookAccessor.SalesDetail.Replace(replacement);

        tested.Tables["sales_detail"].Range.Values.Should().BeEquivalentTo(
            tableBook.Tables["sales_detail"].Range.Values,
            options => options.WithStrictOrdering());
    }

    [Fact]
    public void 生成されたTableのReplaceは手書きPOCOのReplaceと同じセル値を保存します()
    {
        var filePath = temporaryFiles.Copy(BasicStructureExcelFilePath);
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));
        var rowType = generatedAssembly.GeneratedType("SalesDetail");
        ReadTableComparison[] replacement =
        [
            new()
            {
                CustomerId = 10,
                Amount = 1.5,
                Description = "first"
            },
            new()
            {
                CustomerId = 20,
                Amount = 2.5,
                Description = "second"
            }
        ];

        using (var book = generatedAssembly.GeneratedInstance<Workbook>(
                   "BasicStructureBook",
                   filePath))
        {
            dynamic bookAccessor = book;
            dynamic generatedRows = CreateRows(
                rowType,
                replacement.Select(it => (it.CustomerId, it.Amount, it.Description)).ToArray());

            bookAccessor.SalesDetail.Replace(generatedRows);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);
        using var handWrittenBook = Workbook.Open(temporaryFiles.Copy(BasicStructureExcelFilePath));
        handWrittenBook.ReadTable<ReadTableComparison>("sales_detail").Replace(replacement);

        tested.Tables["sales_detail"].Range.Values.Should().BeEquivalentTo(
            handWrittenBook.Tables["sales_detail"].Range.Values,
            options => options.WithStrictOrdering());
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
