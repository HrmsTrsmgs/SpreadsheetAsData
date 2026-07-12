using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Worksheetのテスト : IDisposable
{

    readonly Workbook book;
    readonly Worksheet sheet1;
    readonly Worksheet data;

    public Worksheetのテスト()
    {
        book = Workbook.Open(@"TestData\Book1.xlsx");
        sheet1 = book.Sheets["Sheet1"];
        data = book.Sheets["いろいろなデータ"];
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Nameはシート名を返します()
    {
        sheet1.Name.Should().Be("Sheet1");
    }

    [Fact]
    public void Nameは日本語で指定したシート名も適切に扱います()
    {
        data.Name.Should().Be("いろいろなデータ");
    }

    [Fact]
    public void Bookは所属しているWorkbookを取得します()
    {
        sheet1.Book.Should().BeSameAs(book);
    }

    [Fact]
    public void Cellsはセルの参照文字列を文字列を指定してセル取得します()
    {
        sheet1.Cells["C3"].Reference.Should().Be("C3");
    }

    [Fact]
    public void Cellsはセルの参照文字列に存在しないセル名を指定した時にFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => sheet1.Cells["a1"]
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void Cellsは一文字の列名称の列を取得します()
    {
        sheet1.Cells["A1"].Reference.Should().Be("A1");
        sheet1.Cells["Z1"].Reference.Should().Be("Z1");
    }

    [Fact]
    public void Cellsは二文字の列名称の列を取得します()
    {
        sheet1.Cells["AA1"].Reference.Should().Be("AA1");
        sheet1.Cells["ZZ1"].Reference.Should().Be("ZZ1");
    }

    [Fact]
    public void Cellsは三文字の列名称の列を取得します()
    {
        sheet1.Cells["AAA1"].Reference.Should().Be("AAA1");
        sheet1.Cells["XFD1"].Reference.Should().Be("XFD1");
    }

    [Fact]
    public void Cellsは大きすぎる列名を指定した時にFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => sheet1.Cells["XFE1"]
        ).Should().Throw<FormatException>();
        FluentActions.Invoking(
            () => sheet1.Cells["ZZZ1"]
        ).Should().Throw<FormatException>();
        FluentActions.Invoking(
            () => sheet1.Cells["AAAA1"]
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void Cellsは大きな行番号のセルを取得します()
    {
        sheet1.Cells["A1048576"].Reference.Should().Be("A1048576");
    }

    [Fact]
    public void Cellsは大きすぎる行番号を指定した時にFormatExceptionを投げます()
    {
        FluentActions.Invoking(
            () => sheet1.Cells["A1048577"]
        ).Should().Throw<FormatException>();
    }

    [Fact]
    public void Cellsは空のセルも取得します()
    {
        sheet1.Cells["B1"].Reference.Should().Be("B1");
    }

    [Fact]
    public void Cellsはは同じセルの場合は同じオブジェクトを取得します()
    {
        sheet1.Cells["A1"].Should().BeSameAs(sheet1.Cells["A1"]);
    }

    [Fact]
    public void Cellsはセルの座標を数値で指定してセル取得します()
    {
        sheet1.Cells[3, 3].Reference.Should().Be("C3");
    }

    [Fact]
    public void Cellsは指定方法が違っても同じセルの場合は同じオブジェクトを取得します()
    {
        sheet1.Cells[1, 1].Should().BeSameAs(sheet1.Cells["A1"]);
    }

    [Fact]
    public void Rangeは2引数を指定して範囲を取得します()
    {
        sheet1.Range["A1", "C3"].ToString().Should().Be("A1:C3");
        sheet1.Range["B2", "B2"].ToString().Should().Be("B2:B2");
    }

    [Fact]
    public void Rangeは2引数を指定して同じ範囲を指定した場合に同じセルを返します()
    {
        sheet1.Range["A1", "C3"].Should().BeSameAs(sheet1.Range["A1", "C3"]);
    }
}
