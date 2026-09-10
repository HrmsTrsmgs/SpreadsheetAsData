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

    [Fact(Skip = "ブックからシート名と列・行番号で取得するAPIは仕様レビュー待ち。取得先とセルの同一性を確認してから実装する。")]
    public void WorkbookのCellはシート名と列番号と行番号からWorksheetと同じセルを取得します()
    {
        var fromBook = book.Cell["Sheet2", 6, 33];
        var fromWorksheet = sheet2.Cells["F33"];

        fromBook.Should().BeSameAs(fromWorksheet);
    }

    [Fact(Skip = "ブックからシート名とCellNameで取得するAPIは仕様レビュー待ち。列・行番号指定の確認後にこの取得方法を確認する。")]
    public void WorkbookのCellはシート名とCellNameからWorksheetと同じセルを取得します()
    {
        var fromBook = book.Cell["Sheet2", CellName.Parse("F33")];
        var fromWorksheet = sheet2.Cells["F33"];

        fromBook.Should().BeSameAs(fromWorksheet);
    }

    [Fact]
    public void Cellは存在しないブックスコープの名前を指定した場合に失敗します()
    {
        var action = () => book.Cell["not_found"];

        action.Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void Cellは存在しないワークシートスコープの名前を指定した場合に失敗します()
    {
        var action = () => sheet2.Cell["not_found"];

        action.Should().Throw<KeyNotFoundException>();
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

    [Fact]
    public void WorksheetのCellはA1形式のセル参照へ値を書き込めます()
    {
        var tested = sheet2.Cell["A1"];

        tested.Value = 9.9;

        (tested.Value as object).Should().Be(9.9);
    }

    [Fact]
    public void WorksheetのCellは列番号と行番号からCellsと同じセルを取得します()
    {
        var fromCell = sheet2.Cell[6, 33];
        var fromCells = sheet2.Cells["F33"];

        fromCell.Should().BeSameAs(fromCells);
    }

    [Fact]
    public void WorksheetのCellはCellNameからCellsと同じセルを取得します()
    {
        var fromCell = sheet2.Cell[CellName.Parse("F33")];
        var fromCells = sheet2.Cells["F33"];

        fromCell.Should().BeSameAs(fromCells);
    }

}
