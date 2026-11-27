using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public class TableValueConversionのテスト
{
    [Theory]
    [InlineData((double)int.MinValue, true)]
    [InlineData((double)int.MaxValue, true)]
    [InlineData((double)int.MinValue - 1, false)]
    [InlineData((double)int.MaxValue + 1, false)]
    [InlineData(1.5, false)]
    [InlineData(double.NaN, false)]
    [InlineData(double.PositiveInfinity, false)]
    public void 整数かつintの範囲内の数値だけを変換可能と判定します(double value, bool expected)
    {
        TableValueConversion.CanConvertToInt32(value).Should().Be(expected);
    }
}
