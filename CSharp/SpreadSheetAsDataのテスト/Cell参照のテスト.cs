using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Cell参照のテスト : IDisposable
{
    const string セル名参照API仕様保留理由 =
        "API仕様として先に固定。Greenはセル名参照の実装単位ごとに解除する。";

    readonly Workbook book;
    readonly Worksheet sheet2;

    public Cell参照のテスト()
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
    public void Cellはブックスコープの単一セル名からセルを取得します()
    {
        Cell tested = book.Cell["book_cell"];

        tested.Should().BeSameAs(sheet2.Cells["F33"]);
    }

    [Fact]
    public void Cellはワークシートスコープの単一セル名からセルを取得します()
    {
        Cell tested = sheet2.Cell["cell_name"];

        tested.Should().BeSameAs(sheet2.Cells["E25"]);
    }

    [Fact(Skip = セル名参照API仕様保留理由)]
    public void Cellはブックスコープの複数セル名では失敗します()
    {
        var action = () => book.Cell["book_range"];

        action.Should().Throw<InvalidOperationException>();
    }

    [Fact(Skip = セル名参照API仕様保留理由)]
    public void Cellはワークシートスコープの複数セル名では失敗します()
    {
        var action = () => sheet2.Cell["range_name"];

        action.Should().Throw<InvalidOperationException>();
    }
}
