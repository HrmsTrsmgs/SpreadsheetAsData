using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class Cell書き込みのテスト : IDisposable
{
    readonly TemporaryExcelFiles temporaryFiles = new();

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Valueプロパティは数値を設定すると同じセルから取得できます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));
        var tested = book.Sheets["いろいろなデータ"].Cells["A1"];

        tested.Value = 12.34;

        (tested.Value as object).Should().Be(12.34);
    }

    [Fact(Skip = "読み込みAPIと対になるセル値書き込み機能を実装するときに解除する。")]
    public void Valueプロパティは整数を書き込めます()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A1"].Value = 123;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cells["A1"].Value as object)
            .Should().Be(123d);
    }

    [Fact(Skip = "読み込みAPIと対になるセル値書き込み機能を実装するときに解除する。")]
    public void Valueプロパティは文字列を書き込めます()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A3"].Value = "書き込み";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cells["A3"].Value as object)
            .Should().Be("書き込み");
    }

    [Fact(Skip = "読み込みAPIと対になるセル値書き込み機能を実装するときに解除する。")]
    public void Valueプロパティは真偽値を書き込めます()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A2"].Value = false;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cells["A2"].Value as object)
            .Should().BeOfType<bool>().Which.Should().BeFalse();
    }

    [Fact(Skip = "読み込みAPIと対になるセル値書き込み機能を実装するときに解除する。")]
    public void Valueプロパティにnullを指定すると空白セルとして保存します()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");
        object? blankValue = null;

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A1"].Value = blankValue;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cells["A1"].Value as object)
            .Should().BeOfType<BlankValue>();
    }
}
