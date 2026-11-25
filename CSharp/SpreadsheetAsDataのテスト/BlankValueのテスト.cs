using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public class BlankValueのテスト : IDisposable
{
    readonly Workbook book;
    readonly BlankValue tested;

    public BlankValueのテスト()
    {
        book = Workbook.Open(@"TestData\Book1.xlsx");
        object cellValue = book.Sheets["Sheet1"].Cells["B1"].Value;
        cellValue.Should().BeOfType<BlankValue>();
        tested = (BlankValue)cellValue;
    }

    public void Dispose()
    {
        book.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void 比較すると空白と同じとされます()
    {
        (tested == "").Should().BeTrue();
    }

    [Fact]
    public void 比較するとdoubleの0と同じとされます()
    {
        (tested == .0).Should().BeTrue();
    }

    [Fact]
    public void 文字列として演算すると空白と同じとされます()
    {
        (tested + "A").Should().Be("A");
    }

    [Fact]
    public void 数値として加算すると0と同じとされます()
    {
        (tested + 3).Should().Be(3);
    }

    [Fact]
    public void 数値として乗算すると0と同じとされます()
    {
        (tested * 3).Should().Be(0);
    }

    [Fact]
    public void ToStringで中かっこに囲まれたBlankとなります()
    {
        tested.ToString().Should().Be("{Blank}");
    }
}
