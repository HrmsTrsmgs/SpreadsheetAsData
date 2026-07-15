using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class TableColumnCollectionのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly TableColumnCollection columns;

    public TableColumnCollectionのテスト()
    {
        book = Workbook.Open(TestFilePath);
        columns = book.Tables["テーブル2"].Columns;
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Columnsは名前から列を取得します()
    {
        var tested = columns["数値2"];

        tested.Name.Should().Be("数値2");
    }

    [Fact]
    public void Columnsは0始まりの位置から列を取得します()
    {
        var tested = columns[1];

        tested.Name.Should().Be("数値2");
    }

    [Fact]
    public void Columnsは名前と位置から同じ列オブジェクトを取得します()
    {
        var byName = columns["数値2"];
        var byOrdinal = columns[1];

        byName.Should().BeSameAs(byOrdinal);
    }

    [Fact]
    public void Columnsは存在しない名前を指定した場合に失敗します()
    {
        var action = () => _ = columns["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Columnsは範囲外の位置を指定した場合に失敗します()
    {
        var action = () => _ = columns[4];

        action.Should().Throw<ArgumentOutOfRangeException>();
    }
}
