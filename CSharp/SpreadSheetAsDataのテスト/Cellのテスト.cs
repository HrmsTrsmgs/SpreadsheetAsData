using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Cellのテスト : IDisposable
{
    Worksheet いろいろなデータ;
    Cell a1;
    Cell b1;
    Cell a2;
    Cell b2;
    Cell a3;
    Cell b3;

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
        (a1.Value as object).Should().BeOfType<double>().Which.Should().Be(1.1);
        (b1.Value as object).Should().BeOfType<double>().Which.Should().Be(2.2);
    }

    [Fact]
    public void Valueプロパティはboolの値を取得できます()
    {
        (a2.Value as object).Should().BeOfType<bool>().Which.Should().BeTrue();
        (b2.Value as object).Should().BeOfType<bool>().Which.Should().BeFalse();
    }

    [Fact]
    public void Valueプロパティは文字列の値を取得できます()
    {
        (a3.Value as object).Should().BeOfType<string>().Which.Should().Be("あいうえお");
        (b3.Value as object).Should().BeOfType<string>().Which.Should().Be("かきくけこ");
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
