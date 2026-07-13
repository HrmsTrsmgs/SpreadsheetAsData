using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class CellRangeReferenceのテスト
{
    [Fact]
    public void TryParseはA1形式の範囲参照を変換します()
    {
        CellRangeReference.TryParse("A1:B2", out var tested).Should().BeTrue();

        tested.SheetName.Should().BeNull();
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("B2"));
    }

    [Fact]
    public void TryParseは絶対参照の範囲参照を相対参照へ正規化します()
    {
        CellRangeReference.TryParse("$A$1:$B$2", out var tested).Should().BeTrue();

        tested.SheetName.Should().BeNull();
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("B2"));
    }

    [Fact]
    public void TryParseはシート名付き範囲参照を変換します()
    {
        CellRangeReference.TryParse("Sheet1!A1:B2", out var tested).Should().BeTrue();

        tested.SheetName.Should().Be("Sheet1");
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("B2"));
    }

    [Fact]
    public void TryParseはExcelで使用可能な最大列を含む範囲参照を変換します()
    {
        CellRangeReference.TryParse("A1:XFD1", out var tested).Should().BeTrue();

        tested.SheetName.Should().BeNull();
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("XFD1"));
    }

    [Fact]
    public void TryParseはExcelで使用可能な最大行を含む範囲参照を変換します()
    {
        CellRangeReference.TryParse("A1:A1048576", out var tested).Should().BeTrue();

        tested.SheetName.Should().BeNull();
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("A1048576"));
    }

    [Fact]
    public void TryParseはシート名付き単一セル参照を単一セル範囲として変換します()
    {
        CellRangeReference.TryParse("Sheet1!A1", out var tested).Should().BeTrue();

        tested.SheetName.Should().Be("Sheet1");
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("A1"));
    }

    [Fact]
    public void TryParseはシート名付き絶対セル参照を相対参照の単一セル範囲へ正規化します()
    {
        CellRangeReference.TryParse("Sheet1!$A$1", out var tested).Should().BeTrue();

        tested.SheetName.Should().Be("Sheet1");
        tested.TopLeft.Should().Be(CellName.Parse("A1"));
        tested.BottomRight.Should().Be(CellName.Parse("A1"));
    }

    [Theory]
    [InlineData("abc:def")]
    [InlineData("A$$1:B2")]
    public void TryParseはA1形式でないセルを含む範囲参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference, out _).Should().BeFalse();
    }

    [Theory]
    [InlineData("Sheet1!abc")]
    [InlineData("Sheet1!A$$1")]
    public void TryParseはA1形式でない単一セル範囲参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference, out _).Should().BeFalse();
    }

    [Theory]
    [InlineData("A0:B1")]
    [InlineData("A1:B0")]
    [InlineData("XFE1:XFE2")]
    [InlineData("A1048577:B1048577")]
    public void TryParseはExcelで使用可能な範囲を超えたセル参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference, out _).Should().BeFalse();
    }

    [Theory]
    [InlineData(":B2")]
    [InlineData("A1:")]
    public void TryParseは始点または終点がない範囲参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference, out _).Should().BeFalse();
    }
}
