using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class CellRangeReferenceのテスト
{
    const string A1範囲参照の妥当性検証保留理由 =
        "A1形式とExcelで使用可能な範囲の検証をCellNameへ集約するときに解除する。";

    [Theory(Skip = A1範囲参照の妥当性検証保留理由)]
    [InlineData("abc:def")]
    [InlineData("A$$1:B2")]
    public void TryParseはA1形式でないセルを含む範囲参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference).Should().BeNull();
    }

    [Theory(Skip = A1範囲参照の妥当性検証保留理由)]
    [InlineData("A0:B1")]
    [InlineData("A1:B0")]
    [InlineData("XFE1:XFE2")]
    [InlineData("A1048577:B1048577")]
    public void TryParseはExcelで使用可能な範囲を超えたセル参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference).Should().BeNull();
    }

    [Theory(Skip = A1範囲参照の妥当性検証保留理由)]
    [InlineData(":B2")]
    [InlineData("A1:")]
    public void TryParseは始点または終点がない範囲参照を変換しません(
        string reference)
    {
        CellRangeReference.TryParse(reference).Should().BeNull();
    }
}
