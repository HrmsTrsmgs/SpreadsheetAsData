using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Tableのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly Table table;

    public Tableのテスト()
    {
        book = Workbook.Open(TestFilePath);
        table = book.Tables["テーブル2"];
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void NameはExcelテーブル名を返します()
    {
        table.Name.Should().Be("テーブル2");
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void WorksheetはExcelテーブルが属するワークシートを返します()
    {
        table.Worksheet.Should().BeSameAs(
            book.Sheets["Sheet1"]);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void RangeはExcelテーブル全体の対象範囲を返します()
    {
        table.Range.ToString().Should().Be("B6:E9");
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void ColumnsはExcelテーブルの列を定義順に列挙します()
    {
        table.Columns
            .Select(column => column.Name)
            .Should()
            .Equal("数値", "数値2", "文字列", "真偽値");
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Rowsはすべてのデータ行を列挙します()
    {
        table.Rows.Should().HaveCount(3);
    }

    [Fact(Skip = "Excelテーブル仕様を先行追加しているため、実装対象になったテストから解除する。")]
    public void Rowsはデータ行をワークシート上の順序で列挙します()
    {
        table.Rows
            .Select(row => row.WorksheetRowIndex)
            .Should()
            .Equal(7u, 8u, 9u);
    }
}
