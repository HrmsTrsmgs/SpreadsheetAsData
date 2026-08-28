using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Cell参照のテスト : IDisposable
{
    readonly Workbook book;
    readonly Worksheet sheet2;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public Cell参照のテスト()
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

    [Fact(Skip = "読み込みAPIと対になるCell参照経由の書き込み機能を実装するときに解除する。")]
    public void WorksheetのCellで取得したセルへ値を書き込めます()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cell["A1"].Value = 9.9;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cell["A1"].Value as object)
            .Should().Be(9.9);
    }

    [Fact(Skip = "読み込みAPIと対になるCell参照経由の書き込み機能を実装するときに解除する。")]
    public void BookのCellはブックスコープの単一セル定義名へ値を書き込めます()
    {
        var filePath = temporaryFiles.Copy("定義名.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Cell["book_cell"].Value = "book";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Cell["book_cell"].Value as object)
            .Should().Be("book");
    }

    [Fact(Skip = "読み込みAPIと対になるCell参照経由の書き込み機能を実装するときに解除する。")]
    public void WorksheetのCellはワークシートスコープの単一セル定義名へ値を書き込めます()
    {
        var filePath = temporaryFiles.Copy("定義名.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["Sheet2"].Cell["cell_name"].Value = "sheet";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["Sheet2"].Cell["cell_name"].Value as object)
            .Should().Be("sheet");
    }
}
