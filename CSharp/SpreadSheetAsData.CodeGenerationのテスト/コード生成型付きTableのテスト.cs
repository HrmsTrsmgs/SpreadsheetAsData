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

    [Fact(Skip = "読み込みコード生成と対になる型付きTable書き込み機能を実装するときに解除する。")]
    public void 生成されたTable型のWriteで型付き行を書き込めます()
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

            bookAccessor.SalesDetail.Write(CreateRows(
                rowType,
                (10, 1.5, "first"),
                (20, 2.5, "second")));
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
        [SpreadsheetColumn("customer_id")]
        public int CustomerId { get; set; }

        public double Amount { get; set; }

        public string Description { get; set; } = "";
    }
}
