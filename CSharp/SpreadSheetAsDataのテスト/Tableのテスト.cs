using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Tableのテスト : IDisposable
{
    const string TestFilePath = @"TestData\テーブル.xlsx";

    readonly Workbook book;
    readonly Table table;
    readonly Table typedMappingTable;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public Tableのテスト()
    {
        book = Workbook.Open(TestFilePath);
        table = book.Tables["型付き行マッピング"];
        typedMappingTable = book.Tables["型付き行マッピング"];
    }

    public void Dispose()
    {
        book.Close();
        temporaryFiles.Dispose();
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

    [Fact(Skip = "読み込みAPIと対になるTable単位の型付き書き込み機能を実装するときに解除する。")]
    public void Writeは属性がないプロパティ名を列名として使用します()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Tables["プロパティ名マッピング"].Write(
                [
                    new WritablePropertyNameMappedRow
                    {
                        @float = 10.1,
                        @string = "first"
                    },
                    new WritablePropertyNameMappedRow
                    {
                        @float = 20.2,
                        @string = "second"
                    }
                ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.ReadTable<型付きTableのテスト.PropertyNameMappedRow>("プロパティ名マッピング")
            .Should().BeEquivalentTo(
                [
                    new 型付きTableのテスト.PropertyNameMappedRow
                    {
                        @float = 10.1,
                        @string = "first"
                    },
                    new 型付きTableのテスト.PropertyNameMappedRow
                    {
                        @float = 20.2,
                        @string = "second"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    public sealed class WritablePropertyNameMappedRow
    {
        public double @float { get; set; }

        public string @string { get; set; } = "";
    }
}
