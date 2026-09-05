using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class TableColumnのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly Table table;
    readonly TableColumn column;

    public TableColumnのテスト()
    {
        book = Workbook.Open(TestFilePath);
        table = book.Tables["型付き行マッピング"];
        column = table.Columns["数値2"];
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void NameはExcelテーブルの列名を返します()
    {
        column.Name.Should().Be("数値2");
    }

    [Fact]
    public void ToStringはExcelテーブルの列名を返します()
    {
        column.ToString().Should().Be("数値2");
        table.Columns["文字列"].ToString().Should().Be("文字列");
    }

    [Fact]
    public void Ordinalはテーブル内の0始まりの列位置を返します()
    {
        column.Ordinal.Should().Be(1);
    }

    [Fact]
    public void Tableは列が属するExcelテーブルを返します()
    {
        column.Table.Should().BeSameAs(table);
    }
}
