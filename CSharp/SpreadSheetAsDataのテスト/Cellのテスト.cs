using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Cellのテスト : IDisposable
{
    readonly Worksheet いろいろなデータ;
    readonly Cell a1;
    readonly Cell b1;
    readonly Cell a2;
    readonly Cell b2;
    readonly Cell a3;
    readonly Cell b3;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public Cellのテスト()
    {
        var book = Workbook.Open(@"TestData\Book1.xlsx");

        いろいろなデータ = book.Sheets["いろいろなデータ"];

        a1 = いろいろなデータ.Cells["A1"];
        b1 = いろいろなデータ.Cells["B1"];
        a2 = いろいろなデータ.Cells["A2"];
        b2 = いろいろなデータ.Cells["B2"];
        a3 = いろいろなデータ.Cells["A3"];
        b3 = いろいろなデータ.Cells["B3"];

    }
    public void Dispose()
    {
        いろいろなデータ.Book.Close();
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Bookプロパティはブックを取得できます()
    {
        a1.Book.Should().BeSameAs(いろいろなデータ.Book);
    }

    [Fact]
    public void Sheetプロパティはシートを取得できます()
    {
        a1.Sheet.Should().BeSameAs(いろいろなデータ);
    }

    [Fact]
    public void Valueプロパティは数字の値を取得できます()
    {
        var a1Value = a1.Value as object;
        var b1Value = b1.Value as object;

        a1Value.Should().BeOfType<double>().Which.Should().Be(1.1);
        b1Value.Should().BeOfType<double>().Which.Should().Be(2.2);
    }

    [Fact]
    public void Valueプロパティはboolの値を取得できます()
    {
        var a2Value = a2.Value as object;
        var b2Value = b2.Value as object;

        a2Value.Should().BeOfType<bool>().Which.Should().BeTrue();
        b2Value.Should().BeOfType<bool>().Which.Should().BeFalse();
    }

    [Fact]
    public void Valueプロパティは共有文字列セルの値を取得できます()
    {
        var a3Value = a3.Value as object;
        var b3Value = b3.Value as object;

        a3Value.Should().BeOfType<string>().Which.Should().Be("あいうえお");
        b3Value.Should().BeOfType<string>().Which.Should().Be("かきくけこ");
    }

    [Fact]
    public void Valueプロパティは文字列セルの値を取得できます()
    {
        using var book = Workbook.Open(@"TestData\文字列セル.xlsx");
        var tested = book.Sheets["Sheet1"].Cells["A1"].Value as object;

        tested.Should().Be("直接文字列");
    }

    [Fact(Skip = "読み込みAPIと対になるセル値書き込み機能を実装するときに解除する。")]
    public void Valueプロパティは数値を書き込めます()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A1"].Value = 12.34;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cells["A1"].Value as object)
            .Should().Be(12.34);
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

    [Fact]
    public void Referenceプロパティがセル参照の名称を取得できます()
    {
        a1.Reference.Should().Be("A1");
        b1.Reference.Should().Be("B1");
    }
    [Fact]
    public void RowIndexプロパティが行番号を取得できます()
    {
        a1.RowIndex.Should().Be(1U);
        a2.RowIndex.Should().Be(2U);
    }

    [Fact]
    public void ColumnIndexプロパティが列番号を取得できます()
    {
        a1.ColumnIndex.Should().Be(1U);
        b1.ColumnIndex.Should().Be(2U);
    }
}
