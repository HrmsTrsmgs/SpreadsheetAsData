using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

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
    public void Nameはname属性がない場合displayNameを返します()
    {
        using var book = Workbook.Open(@"TestData\テーブル名属性省略.xlsx");

        book.Tables["型付き行マッピング"].Name.Should().Be("型付き行マッピング");
    }

    [Fact]
    public void Nameは両方の名前属性がない場合InvalidDataExceptionを投げます()
    {
        using var book = Workbook.Open(@"TestData\テーブル名属性全省略.xlsx");
        var tested = book.Tables.Single(it => it.Range.ToString() == "B6:E9");
        var action = () => { _ = tested.Name; };

        action.Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void ToStringはExcelテーブル名を返します()
    {
        table.ToString().Should().Be("型付き行マッピング");
        book.Tables["別シート行列挙"].ToString().Should().Be("別シート行列挙");
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
    public void 範囲属性のないExcelテーブルはInvalidDataExceptionで拒否します()
    {
        using var book = Workbook.Open(@"TestData\テーブル範囲属性省略.xlsx");

        FluentActions.Invoking(() => book.Tables["型付き行マッピング"])
            .Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void テーブル要素がない場合InvalidDataExceptionで拒否します()
    {
        using var book = Workbook.Open(@"TestData\テーブル要素省略.xlsx");

        FluentActions.Invoking(() => book.Tables["型付き行マッピング"])
            .Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void ColumnsはExcelテーブルの列を定義順に列挙します()
    {
        (
            from column in table.Columns
            select column.Name
        )
            .Should().Equal("数値", "数値2", "文字列", "真偽値");
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
    public void セル参照属性のないセルを含むExcelテーブルは拒否します()
    {
        using var book = Workbook.Open(@"TestData\テーブル内セル参照属性省略.xlsx");

        FluentActions.Invoking(() => book.Tables["型付き行マッピング"])
            .Should().Throw<NotSupportedException>();
    }

    [Fact]
    public void Rowsはデータ行をワークシート上の順序で列挙します()
    {
        (
            from row in table.Rows
            select row.WorksheetRowIndex
        )
            .Should().Equal(7u, 8u, 9u);
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
            .Should().BeEquivalentTo(
                book.ReadTable<TestMappedRow>("型付き行マッピング"),
                options => options.WithStrictOrdering());
    }

}
