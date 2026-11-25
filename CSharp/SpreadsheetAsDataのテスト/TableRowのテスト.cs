using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Marimo.SpreadsheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public class TableRowのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly Table table;
    readonly TableRow tested;
    readonly TableColumn secondColumn;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public TableRowのテスト()
    {
        book = Workbook.Open(TestFilePath);
        table = book.Tables["型付き行マッピング"];
        tested = table.Rows.First();
        secondColumn = table.Columns["数値2"];
    }

    public void Dispose()
    {
        book.Close();
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Ordinalはテーブル内の0始まりのデータ行位置を返します()
    {
        tested.Ordinal.Should().Be(0);
    }

    [Fact]
    public void WorksheetRowIndexはワークシート上の1始まりの行番号を返します()
    {
        tested.WorksheetRowIndex.Should().Be(7u);
    }

    [Fact]
    public void Tableは行が属するExcelテーブルを返します()
    {
        tested.Table.Should().BeSameAs(table);
    }

    [Fact]
    public void 列名を指定すると対応するセルを返します()
    {
        tested["数値2"]
            .Reference.Should().Be("C7");
    }

    [Fact]
    public void TableRowの列名とWorksheetのセル参照から取得した同じセルは同一オブジェクトです()
    {
        var fromTableRow = tested["数値2"];
        var fromWorksheet = table.Worksheet.Cells["C7"];

        fromTableRow.Should().BeSameAs(fromWorksheet);
    }

    [Fact]
    public void TableColumnを指定すると対応するセルを返します()
    {
        tested[secondColumn]
            .Reference.Should().Be("C7");
    }

    [Fact]
    public void TableRowのTableColumnとWorksheetのセル参照から取得した同じセルは同一オブジェクトです()
    {
        var fromTableRow = tested[secondColumn];
        var fromWorksheet = table.Worksheet.Cells["C7"];

        fromTableRow.Should().BeSameAs(fromWorksheet);
    }

    [Fact]
    public void 列位置を指定すると対応するセルを返します()
    {
        tested[1]
            .Reference.Should().Be("C7");
    }

    [Fact]
    public void TableRowの列位置とWorksheetのセル参照から取得した同じセルは同一オブジェクトです()
    {
        var fromTableRow = tested[1];
        var fromWorksheet = table.Worksheet.Cells["C7"];

        fromTableRow.Should().BeSameAs(fromWorksheet);
    }

    [Fact]
    public void 存在しない列名を指定した場合に失敗します()
    {
        var action = () => _ = tested["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void 別のExcelテーブルに属するTableColumnを指定した場合に失敗します()
    {
        var foreignColumn =
            book.Tables["別テーブル列検証"]
                .Columns["数値2"];

        var action = () => _ = tested[foreignColumn];

        action.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void 範囲外の列位置を指定した場合に失敗します()
    {
        var action = () => _ = tested[4];

        action.Should().Throw<ArgumentOutOfRangeException>();
    }

    [Fact]
    public void 列名を指定して対応するセルへ値を書き込めます()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Tables["型付き行マッピング"].Rows.First()["数値2"].Value = 99;
            book.Save();
        }

        using var savedBook = Workbook.Open(filePath);

        (savedBook.Tables["型付き行マッピング"].Rows.First()["数値2"].Value as object)
            .Should().Be(99d);
    }
}
