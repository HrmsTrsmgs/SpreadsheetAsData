using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Cell参照のテスト : IDisposable
{
    readonly Workbook book;
    readonly Worksheet sheet2;

    public Cell参照のテスト()
    {
        book = Workbook.Open(@"TestData\定義名.xlsx");
        sheet2 = book.Sheets["Sheet2"];
    }

    public void Dispose()
    {
        book.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Cellはブックスコープの単一セル名からセルを取得します()
    {
        var tested = book.Cell["book_cell"];

        tested.Should().BeSameAs(sheet2.Cells["F33"]);
    }

    [Fact]
    public void Cellはワークシートスコープの単一セル名からセルを取得します()
    {
        var tested = sheet2.Cell["cell_name"];

        tested.Should().BeSameAs(sheet2.Cells["E25"]);
    }

    [Fact]
    public void Cellはブックスコープの複数セル名では失敗します()
    {
        var action = () => book.Cell["book_range"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact]
    public void Cellはワークシートスコープの複数セル名では失敗します()
    {
        var action = () => sheet2.Cell["range_name"];

        action.Should().Throw<InvalidOperationException>();
    }
}
