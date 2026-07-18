using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class RangeExtensionsのテスト
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Containsは開始位置以上で終了位置より小さい値を含みます(int value)
    {
        (0..3).Contains(value).Should().BeTrue();
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(3)]
    public void Containsは開始位置未満または終了位置以上の値を含みません(int value)
    {
        (0..3).Contains(value).Should().BeFalse();
    }
}
