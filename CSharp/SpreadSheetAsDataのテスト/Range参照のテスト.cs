using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Range参照のテスト : IDisposable
{
    readonly Workbook book;
    readonly Worksheet sheet2;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public Range参照のテスト()
    {
        book = Workbook.Open(@"TestData\定義名.xlsx");
        sheet2 = book.Sheets["Sheet2"];
    }

    public void Dispose()
    {
        book.Dispose();
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void RangeはA1形式の範囲参照から名前なし範囲を取得します()
    {
        var tested = sheet2.Range["C32:D36"];

        tested.Should().BeSameAs(sheet2.Range["C32", "D36"]);
    }

    [Fact]
    public void Rangeはブックスコープの名前参照から名前付き範囲を取得します()
    {
        var tested = book.Range["book_range"];

        tested.TopLeftCell.Should().BeSameAs(sheet2.Cells["C32"]);
        tested.BottomRightCell.Should().BeSameAs(sheet2.Cells["D36"]);
    }

    [Fact]
    public void RangeはA1形式のセル参照と同じ名前の定義名を優先します()
    {
        var tested = book.Range["A1"];

        tested.Name.Should().Be("A1");
        tested.TopLeftCell.Should().BeSameAs(sheet2.Cells["B23"]);
        tested.BottomRightCell.Should().BeSameAs(sheet2.Cells["C27"]);
    }

    [Fact]
    public void Rangeはブックスコープの単一セル名を名前付き範囲として取得します()
    {
        var tested = book.Range["book_cell"];

        tested.TopLeftCell.Should().BeSameAs(sheet2.Cells["F33"]);
        tested.BottomRightCell.Should().BeSameAs(sheet2.Cells["F33"]);
    }

    [Fact]
    public void Rangeはワークシートスコープの名前参照から名前付き範囲を取得します()
    {
        var range = sheet2.Range["range_name"];
        var cell = sheet2.Range["cell_name"];

        range.TopLeftCell.Should().BeSameAs(sheet2.Cells["B23"]);
        range.BottomRightCell.Should().BeSameAs(sheet2.Cells["C27"]);
        cell.TopLeftCell.Should().BeSameAs(sheet2.Cells["E25"]);
        cell.BottomRightCell.Should().BeSameAs(sheet2.Cells["E25"]);
    }

    [Fact]
    public void Rangeは同じ名前参照から取得した範囲を同一オブジェクトとして扱います()
    {
        book.Range["book_range"].Should().BeSameAs(book.Range["book_range"]);
        sheet2.Range["range_name"].Should().BeSameAs(sheet2.Range["range_name"]);
    }

    [Fact]
    public void Rangeは同じA1形式の範囲参照から取得した範囲を同一オブジェクトとして扱います()
    {
        sheet2.Range["C32:D36"].Should().BeSameAs(sheet2.Range["C32:D36"]);
    }

    [Fact]
    public void Rangeは名前参照と同じ位置を指すA1形式の範囲参照を別オブジェクトとして扱います()
    {
        book.Range["book_range"].Should().NotBeSameAs(sheet2.Range["C32", "D36"]);
        sheet2.Range["range_name"].Should().NotBeSameAs(sheet2.Range["B23", "C27"]);
    }

    [Fact]
    public void NameはA1形式の範囲参照から取得した範囲ではnullを返します()
    {
        var tested = sheet2.Range["C32:D36"];

        tested.Name.Should().BeNull();
    }

    [Fact]
    public void Nameはブックスコープの名前参照から取得した範囲では名前を返します()
    {
        var range = book.Range["book_range"];
        var cell = book.Range["book_cell"];

        range.Name.Should().Be("book_range");
        cell.Name.Should().Be("book_cell");
    }

    [Fact]
    public void Nameはワークシートスコープの名前参照から取得した範囲では名前を返します()
    {
        var range = sheet2.Range["range_name"];
        var cell = sheet2.Range["cell_name"];

        range.Name.Should().Be("range_name");
        cell.Name.Should().Be("cell_name");
    }

    [Fact]
    public void ToStringは名前なし範囲ではA1形式の範囲参照を返します()
    {
        var tested = sheet2.Range["C32", "D36"];

        tested.ToString().Should().Be("C32:D36");
    }

    [Fact]
    public void ToStringは名前付き範囲では名前を返します()
    {
        var tested = book.Range["book_range"];

        tested.ToString().Should().Be("book_range");
    }

    [Fact]
    public void Valuesは範囲内のセル値を行ごとに列挙します()
    {
        using var book = Workbook.Open(@"TestData\Book1.xlsx");

        var rows = book.Sheets["いろいろなデータ"].Range["A1:B2"].Values
            .Select(it => it.ToArray())
            .ToArray();

        rows.Should().HaveCount(2);
        rows[0].Should().Equal(1.1, 2.2);
        rows[1].Should().Equal(true, false);
    }

    [Fact]
    public void ValuesはA1形式の範囲へ行ごとの値を書き込めます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));
        var sheet = book.Sheets["いろいろなデータ"];

        sheet.Range["A1:B2"].Values =
        [
            [1, "a"],
            [true, null]
        ];

        (sheet.Cells["A1"].Value as object).Should().Be(1d);
        (sheet.Cells["B1"].Value as object).Should().Be("a");
        (sheet.Cells["A2"].Value as object)
            .Should().BeOfType<bool>().Which.Should().BeTrue();
        (sheet.Cells["B2"].Value as object).Should().BeOfType<BlankValue>();
    }

    [Fact]
    public void Valuesは範囲サイズと値サイズが違う場合に失敗します()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));

        var action = () =>
        {
            tested.Sheets["いろいろなデータ"].Range["A1:B2"].Values =
            [
                ["a"]
            ];
        };

        action.Should().Throw<ArgumentException>();
    }
}
