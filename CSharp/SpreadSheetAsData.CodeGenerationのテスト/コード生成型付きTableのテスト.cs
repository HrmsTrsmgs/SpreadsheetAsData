using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型付きTableのテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";

    [Fact]
    public void 生成されたTableはPOCOをExcel上の順序で列挙します()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        var rows = ((IEnumerable<object>)((dynamic)book).SalesDetail)
            .ToArray();

        rows.Select(it => PropertyValue(it, "CustomerId"))
            .Should().Equal(1, 2);
        rows.Select(it => PropertyValue(it, "Amount"))
            .Should().Equal(10.5, 20.5);
        rows.Select(it => PropertyValue(it, "Description"))
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

        Table tested = ((dynamic)book).SalesDetail;

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

        Table tested = ((dynamic)book).SalesDetail;

        tested.Name.Should().Be("SalesDetail");
        tested.Worksheet.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested.Columns.Should().NotBeEmpty();
    }

    [Fact]
    public void 生成された行データ型は利用者定義POCOと同じ変換規則で読み込まれます()
    {
        (object? CustomerId, object? Amount, object? Description)[] generatedRows;

        using (var generatedBook = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           BasicStructureExcelFilePath))
                   .GeneratedInstance<Workbook>("BasicStructureBook"))
        {
            generatedRows = ((IEnumerable<object>)((dynamic)generatedBook).SalesDetail)
                .Select(ReadGeneratedRow)
                .ToArray();
        }

        using var book = Workbook.Open(BasicStructureExcelFilePath);

        generatedRows.Should().Equal(
                book.ReadTable<ReadTableComparison>("SalesDetail")
                    .Select(ReadHandWrittenRow));
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

    sealed class ReadTableComparison
    {
        public int CustomerId { get; set; }

        public double Amount { get; set; }

        public string Description { get; set; } = "";
    }
}
