using System.Collections;
using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型付きTableのテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\基本構造.xlsx";

    [Fact(
        Skip =
            "生成TableがPOCOをExcel上の順序で列挙する処理を実装するときに解除する。")]
    public void 生成されたTableはPOCOをExcel上の順序で列挙します()
    {
        var table = Activator.CreateInstance(
            CodeGenerationSpec
                .CompileGeneratedAssembly(BasicStructureExcelFilePath)
                .GetRequiredType("SalesDetailTable"));

        var rows = ((IEnumerable)table!).Cast<object>().ToArray();

        rows.Select(it => it.GetType().GetProperty("CustomerId")!.GetValue(it))
            .Should()
            .Equal(1, 2);
        rows.Select(it => it.GetType().GetProperty("Amount")!.GetValue(it))
            .Should()
            .Equal(10.5, 20.5);
        rows.Select(it => it.GetType().GetProperty("Description")!.GetValue(it))
            .Should()
            .Equal("a", "b");
    }

    [Fact(
        Skip =
            "生成TableをTableとして扱った場合に非型付きRowsを利用できる継承構造を実装するときに解除する。")]
    public void 生成されたTableをTableとして扱うと非型付き行を利用できます()
    {
        var table = (Table)Activator.CreateInstance(
            CodeGenerationSpec
                .CompileGeneratedAssembly(BasicStructureExcelFilePath)
                .GetRequiredType("SalesDetailTable"))!;

        table.Rows.Should().NotBeEmpty();
    }

    [Fact(
        Skip =
            "生成TableからTableの構造情報を利用できる継承構造を実装するときに解除する。")]
    public void 生成されたTable型からTableの構造情報を使用できます()
    {
        var table = (Table)Activator.CreateInstance(
            CodeGenerationSpec
                .CompileGeneratedAssembly(BasicStructureExcelFilePath)
                .GetRequiredType("SalesDetailTable"))!;

        table.Name.Should().Be("sales_detail");
        table.Worksheet.Should().NotBeNull();
        table.Range.Should().NotBeNull();
        table.Columns.Should().NotBeEmpty();
    }

    [Fact(
        Skip =
            "生成された行データ型がReadTableの利用者定義POCOと同じ変換規則で読み込まれる処理を実装するときに解除する。")]
    public void 生成された行データ型は利用者定義POCOと同じ変換規則で読み込まれます()
    {
        using var book = Workbook.Open(BasicStructureExcelFilePath);
        var generatedTable = (IEnumerable)Activator.CreateInstance(
            CodeGenerationSpec
                .CompileGeneratedAssembly(BasicStructureExcelFilePath)
                .GetRequiredType("SalesDetailTable"))!;

        generatedTable
            .Cast<object>()
            .Select(ReadGeneratedRow)
            .Should()
            .Equal(
                book.ReadTable<ReadTableComparison>("sales_detail")
                    .Select(ReadHandWrittenRow));
    }

    static object[] ReadGeneratedRow(object row) =>
        [
            row.GetType().GetProperty("CustomerId")!.GetValue(row)!,
            row.GetType().GetProperty("Amount")!.GetValue(row)!,
            row.GetType().GetProperty("Description")!.GetValue(row)!
        ];

    static object[] ReadHandWrittenRow(ReadTableComparison row) =>
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


