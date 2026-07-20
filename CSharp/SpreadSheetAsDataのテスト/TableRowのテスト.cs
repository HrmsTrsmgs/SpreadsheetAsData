using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class TableRowのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly Table table;
    readonly TableRow firstRow;
    readonly TableColumn secondColumn;

    public TableRowのテスト()
    {
        book = Workbook.Open(TestFilePath);
        table = book.Tables["型付き行マッピング"];
        firstRow = table.Rows.First();
        secondColumn = table.Columns["数値2"];
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Ordinalはテーブル内の0始まりのデータ行位置を返します()
    {
        firstRow.Ordinal.Should().Be(0);
    }

    [Fact]
    public void WorksheetRowIndexはワークシート上の1始まりの行番号を返します()
    {
        firstRow.WorksheetRowIndex.Should().Be(7u);
    }

    [Fact]
    public void Tableは行が属するExcelテーブルを返します()
    {
        firstRow.Table.Should().BeSameAs(table);
    }

    [Fact]
    public void 列名を指定すると対応するセルを返します()
    {
        firstRow["数値2"]
            .Reference.Should().Be("C7");
    }

    [Fact]
    public void TableColumnを指定すると対応するセルを返します()
    {
        firstRow[secondColumn]
            .Reference.Should().Be("C7");
    }

    [Fact]
    public void 列位置を指定すると対応するセルを返します()
    {
        firstRow[1]
            .Reference.Should().Be("C7");
    }

    [Fact]
    public void 存在しない列名を指定した場合に失敗します()
    {
        var action = () => _ = firstRow["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void 別のExcelテーブルに属するTableColumnを指定した場合に失敗します()
    {
        var foreignColumn =
            book.Tables["別テーブル列検証"]
                .Columns["数値2"];

        var action = () => _ = firstRow[foreignColumn];

        action.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void 範囲外の列位置を指定した場合に失敗します()
    {
        var action = () => _ = firstRow[4];

        action.Should().Throw<ArgumentOutOfRangeException>();
    }
}
