using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using System.Globalization;
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

    [Fact]
    public void Valueプロパティは整数を設定すると同じセルから数値として取得できます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));
        var tested = book.Sheets["いろいろなデータ"].Cells["A1"];

        tested.Value = 123;

        (tested.Value as object).Should().Be(123d);
    }

    [Fact]
    public void Valueプロパティは現在カルチャーに依存せず数値を書き込めます()
    {
        var originalCulture = CultureInfo.CurrentCulture;

        try
        {
            using var book = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));
            var tested = book.Sheets["いろいろなデータ"].Cells["A1"];

            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");

            tested.Value = 12.34;

            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;

            (tested.Value as object).Should().Be(12.34);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    [Fact]
    public void Valueプロパティは文字列を書き込めます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));
        var tested = book.Sheets["いろいろなデータ"].Cells["A3"];

        tested.Value = "書き込み";

        (tested.Value as object).Should().Be("書き込み");
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
