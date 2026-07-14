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

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Tablesは名前からExcelテーブルを取得します()
    {
        var tested = book.Tables["テーブル2"];

        tested.Name.Should().Be("テーブル2");
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Tablesは列挙したExcelテーブルと名前から取得したExcelテーブルに同じオブジェクトを返します()
    {
        var enumerated = book.Tables.Single(
            table => table.Name == "テーブル2");

        var tested = book.Tables["テーブル2"];

        tested.Should().BeSameAs(enumerated);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Tablesは異なるワークシートにあるExcelテーブルも列挙します()
    {
        var tested = book.Tables.Single(
            table => table.Name == "テーブル6");

        tested.Worksheet.Should().BeSameAs(book.Sheets["Sheet3"]);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Tablesは存在しない名前を指定した場合に失敗します()
    {
        var action = () => _ = book.Tables["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }
}
