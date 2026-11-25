using FluentAssertions;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public class EnumerableExtensionsのテスト
{
    static readonly string[] values = ["a", "b", "c"];

    [Fact]
    public void WithIndexは各要素に0始まりの位置を付けます()
    {
        values.WithIndex().Should().Equal(("a", 0), ("b", 1), ("c", 2));
    }

    [Fact]
    public void WithIndexは指定した開始位置から各要素に位置を付けます()
    {
        values.WithIndex(1).Should().Equal(("a", 1), ("b", 2), ("c", 3));
    }
}
