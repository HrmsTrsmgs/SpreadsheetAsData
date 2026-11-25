using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public class TableColumnCollectionのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly TableColumnCollection tested;

    public TableColumnCollectionのテスト()
    {
        book = Workbook.Open(TestFilePath);
        tested = book.Tables["型付き行マッピング"].Columns;
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Columnsは名前から列を取得します()
    {
        var byName = tested["数値2"];

        byName.Name.Should().Be("数値2");
    }

    [Fact]
    public void Columnsは0始まりの位置から列を取得します()
    {
        var byOrdinal = tested[1];

        byOrdinal.Name.Should().Be("数値2");
    }

    [Fact]
    public void Columnsは名前と位置から同じ列オブジェクトを取得します()
    {
        var byName = tested["数値2"];
        var byOrdinal = tested[1];

        byName.Should().BeSameAs(byOrdinal);
    }

    [Fact]
    public void Columnsは存在しない名前を指定した場合に失敗します()
    {
        var action = () => _ = tested["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void Containsは指定した名前の列が存在するかを返します()
    {
        tested.Contains("数値2").Should().BeTrue();
        tested.Contains("not_found").Should().BeFalse();
    }

    [Fact]
    public void CountはExcelテーブルの列数を返します()
    {
        tested.Count.Should().Be(4);
    }

    [Fact]
    public void 列定義のないExcelテーブルはInvalidDataExceptionで拒否します()
    {
        using var book = Workbook.Open(@"TestData\テーブル列定義省略.xlsx");

        FluentActions.Invoking(() => book.Tables["型付き行マッピング"])
            .Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void Columnsは範囲外の位置を指定した場合に失敗します()
    {
        var action = () => _ = tested[4];

        action.Should().Throw<ArgumentOutOfRangeException>();
    }
}
