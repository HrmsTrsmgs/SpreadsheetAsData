using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class TableCollectionのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;

    public TableCollectionのテスト()
    {
        book = Workbook.Open(TestFilePath);
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Tablesはブック内のExcelテーブルを列挙します()
    {
        book.Tables.Should().HaveCount(8);
    }

    [Fact]
    public void Tablesは名前からExcelテーブルを取得します()
    {
        var tested = book.Tables["型付き行マッピング"];

        tested.Name.Should().Be("型付き行マッピング");
    }

    [Fact]
    public void Tablesは列挙したExcelテーブルと名前から取得したExcelテーブルに同じオブジェクトを返します()
    {
        var enumerated = (
            from table in book.Tables
            where table.Name == "型付き行マッピング"
            select table
        ).Single();

        var tested = book.Tables["型付き行マッピング"];

        tested.Should().BeSameAs(enumerated);
    }

    [Fact]
    public void Tablesは異なるワークシートにあるExcelテーブルも列挙します()
    {
        var tested = (
            from table in book.Tables
            where table.Name == "別シート行列挙"
            select table
        ).Single();

        tested.Worksheet.Should().BeSameAs(book.Sheets["Sheet3"]);
    }

    [Fact]
    public void Tablesは存在しない名前を指定した場合に失敗します()
    {
        var action = () => _ = book.Tables["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }
}
