using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class SpreadSheetNameAttributeのテスト
{
    [Fact]
    public void Nameはコンストラクターで指定したExcel上の名前を返します()
    {
        new SpreadSheetNameAttribute("name")
            .Name
            .Should().Be("name");
    }

    [Fact]
    public void コンストラクターはnullの名前を拒否します()
    {
        var action = () =>
        {
            _ = new SpreadSheetNameAttribute(null!);
        };

        action
            .Should().Throw<ArgumentNullException>()
            .WithParameterName("name");
    }

    [Fact]
    public void コンストラクターは空の名前を拒否します()
    {
        var action = () =>
        {
            _ = new SpreadSheetNameAttribute("");
        };

        action
            .Should().Throw<ArgumentException>()
            .WithParameterName("name");
    }
}
