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
        table = book.Tables["テーブル2"];
        column = table.Columns["数値2"];
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void NameはExcelテーブルの列名を返します()
    {
        column.Name.Should().Be("数値2");
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Ordinalはテーブル内の0始まりの列位置を返します()
    {
        column.Ordinal.Should().Be(1);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Tableは列が属するExcelテーブルを返します()
    {
        column.Table.Should().BeSameAs(table);
    }
}
