using FluentAssertions;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class EnumerableExtensionsのテスト
{
    [Fact]
    public void WithIndexは各要素に0始まりの位置を付けます()
    {
        var tested = new[] { "a", "b", "c" }
            .WithIndex()
            .ToArray();

        tested
            .Select(item => item.Value)
            .Should()
            .Equal("a", "b", "c");

        tested
            .Select(item => item.Index)
            .Should()
            .Equal(0, 1, 2);
    }

    [Fact]
    public void WithIndexは指定した開始位置から各要素に位置を付けます()
    {
        var tested = new[] { "a", "b", "c" }
            .WithIndex(1)
            .ToArray();

        tested
            .Select(item => item.Value)
            .Should()
            .Equal("a", "b", "c");

        tested
            .Select(item => item.Index)
            .Should()
            .Equal(1, 2, 3);
    }
}
