using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Range参照のテスト : IDisposable
{
    const string 範囲参照API仕様保留理由 = "API仕様として先に固定。Greenは範囲参照の実装単位ごとに解除する。";

    readonly Workbook book;
    readonly Worksheet sheet2;

    public Range参照のテスト()
    {
        book = Workbook.Open(@"TestData\テーブル.xlsx");
        sheet2 = book.Sheets["Sheet2"];
    }

    public void Dispose()
    {
        book.Dispose();
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
        CellRange tested = book.Range["book_range"];

        tested.TopLeftCell.Should().BeSameAs(sheet2.Cells["C32"]);
        tested.BottomRightCell.Should().BeSameAs(sheet2.Cells["D36"]);
    }

    [Fact]
    public void Rangeはブックスコープの単一セル名を名前付き範囲として取得します()
    {
        CellRange tested = book.Range["book_cell"];

        tested.TopLeftCell.Should().BeSameAs(sheet2.Cells["F33"]);
        tested.BottomRightCell.Should().BeSameAs(sheet2.Cells["F33"]);
    }

    [Fact]
    public void Rangeはワークシートスコープの名前参照から名前付き範囲を取得します()
    {
        CellRange range = sheet2.Range["range_name"];
        CellRange cell = sheet2.Range["cell_name"];

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

    [Fact(Skip = 範囲参照API仕様保留理由)]
    public void Rangeは同じA1形式の範囲参照から取得した範囲を同一オブジェクトとして扱います()
    {
        sheet2.Range["C32:D36"].Should().BeSameAs(sheet2.Range["C32:D36"]);
    }

    [Fact(Skip = 範囲参照API仕様保留理由)]
    public void Rangeは名前参照と同じ位置を指すA1形式の範囲参照を別オブジェクトとして扱います()
    {
        book.Range["book_range"].Should().NotBeSameAs(sheet2.Range["C32", "D36"]);
        sheet2.Range["range_name"].Should().NotBeSameAs(sheet2.Range["B23", "C27"]);
    }

    [Fact(Skip = 範囲参照API仕様保留理由)]
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

    [Fact(Skip = 範囲参照API仕様保留理由)]
    public void Nameはワークシートスコープの名前参照から取得した範囲では名前を返します()
    {
        var range = sheet2.Range["range_name"];
        var cell = sheet2.Range["cell_name"];

        range.Name.Should().Be("range_name");
        cell.Name.Should().Be("cell_name");
    }

    [Fact(Skip = 範囲参照API仕様保留理由)]
    public void ToStringは名前なし範囲ではA1形式の範囲参照を返します()
    {
        var tested = sheet2.Range["C32", "D36"];

        tested.ToString().Should().Be("C32:D36");
    }

    [Fact(Skip = 範囲参照API仕様保留理由)]
    public void ToStringは名前付き範囲では名前を返します()
    {
        var tested = book.Range["book_range"];

        tested.ToString().Should().Be("book_range");
    }
}
