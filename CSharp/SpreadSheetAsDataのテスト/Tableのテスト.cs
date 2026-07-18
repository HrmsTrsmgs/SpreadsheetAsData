using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Tableのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly Table table;
    readonly Table typedMappingTable;

    public Tableのテスト()
    {
        book = Workbook.Open(TestFilePath);
        table = book.Tables["型付き行マッピング"];
        typedMappingTable = book.Tables["型付き行マッピング"];
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void NameはExcelテーブル名を返します()
    {
        table.Name.Should().Be("型付き行マッピング");
    }

    [Fact]
    public void WorksheetはExcelテーブルが属するワークシートを返します()
    {
        table.Worksheet.Should().BeSameAs(
            book.Sheets["Sheet1"]);
    }

    [Fact]
    public void RangeはExcelテーブル全体の対象範囲を返します()
    {
        table.Range.ToString().Should().Be("B6:E9");
    }

    [Fact]
    public void ColumnsはExcelテーブルの列を定義順に列挙します()
    {
        (
            from column in table.Columns
            select column.Name
        )
            .Should()
            .Equal("数値", "数値2", "文字列", "真偽値");
    }

    [Fact]
    public void Columnsは同じExcelテーブル列を同一オブジェクトとして扱います()
    {
        var firstEnumeration = table.Columns.ToArray();
        var secondEnumeration = table.Columns.ToArray();

        firstEnumeration[1].Should().BeSameAs(secondEnumeration[1]);
    }

    [Fact]
    public void Rowsはすべてのデータ行を列挙します()
    {
        table.Rows.Should().HaveCount(3);
    }

    [Fact]
    public void Rowsはデータ行をワークシート上の順序で列挙します()
    {
        (
            from row in table.Rows
            select row.WorksheetRowIndex
        )
            .Should()
            .Equal(7u, 8u, 9u);
    }

    [Fact]
    public void Rowsは同じデータ行を同一オブジェクトとして扱います()
    {
        var firstEnumeration = table.Rows.ToArray();
        var secondEnumeration = table.Rows.ToArray();

        firstEnumeration[1].Should().BeSameAs(secondEnumeration[1]);
    }

    [Fact]
    public void EnumerateはReadTableで取得した型付きTableと同じ結果を列挙します()
    {
        typedMappingTable.Enumerate<TestMappedRow>()
            .Should()
            .BeEquivalentTo(
                book.ReadTable<TestMappedRow>("型付き行マッピング"),
                options => options.WithStrictOrdering());
    }
}
