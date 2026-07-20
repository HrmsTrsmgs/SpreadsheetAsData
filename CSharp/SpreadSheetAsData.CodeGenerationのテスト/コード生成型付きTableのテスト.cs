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

    [Fact(
        Skip =
            "生成TableをTableとして扱った場合に非型付きRowsを利用できる継承構造を実装するときに解除する。")]
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

    [Fact(
        Skip =
            "生成TableからTableの構造情報を利用できる継承構造を実装するときに解除する。")]
    public void 生成されたTable型からTableの構造情報を使用できます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        Table tested = ((dynamic)book).SalesDetail;

        tested.Name.Should().Be("sales_detail");
        tested.Worksheet.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested.Columns.Should().NotBeEmpty();
    }

    [Fact(
        Skip =
            "生成された行データ型がReadTableの利用者定義POCOと同じ変換規則で読み込まれる処理を実装するときに解除する。")]
    public void 生成された行データ型は利用者定義POCOと同じ変換規則で読み込まれます()
    {
        using var book = Workbook.Open(BasicStructureExcelFilePath);
        using var generatedBook = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        ((IEnumerable<object>)((dynamic)generatedBook).SalesDetail)
            .Select(ReadGeneratedRow)
            .Should().Equal(
                book.ReadTable<ReadTableComparison>("sales_detail")
                    .Select(ReadHandWrittenRow));
    }

    static object?[] ReadGeneratedRow(object row) =>
        [
            PropertyValue(row, "CustomerId"),
            PropertyValue(row, "Amount"),
            PropertyValue(row, "Description")
        ];

    static object? PropertyValue(object source, string propertyName) =>
        source.GetType()
            .GetProperty(propertyName)
            ?.GetValue(source);

    static object?[] ReadHandWrittenRow(ReadTableComparison row) =>
        [
            row.CustomerId,
            row.Amount,
            row.Description
        ];

    sealed class ReadTableComparison
    {
        [SpreadsheetColumn("customer_id")]
        public int CustomerId { get; set; }

        [SpreadsheetColumn("amount")]
        public double Amount { get; set; }

        [SpreadsheetColumn("description")]
        public string Description { get; set; } = "";
    }
}
