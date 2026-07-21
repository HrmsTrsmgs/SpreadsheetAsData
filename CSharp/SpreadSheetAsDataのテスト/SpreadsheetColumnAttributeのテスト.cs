using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class SpreadsheetColumnAttributeのテスト
{
    [Fact]
    public void Nameはコンストラクターで指定した列名を返します()
    {
        new SpreadsheetColumnAttribute("int")
            .Name
            .Should().Be("int");
    }

    [Fact]
    public void コンストラクターはnullの列名を拒否します()
    {
        var action = () =>
        {
            _ = new SpreadsheetColumnAttribute(null!);
        };

        action
            .Should().Throw<ArgumentNullException>()
            .WithParameterName("name");
    }

    [Fact]
    public void コンストラクターは空の列名を拒否します()
    {
        var action = () =>
        {
            _ = new SpreadsheetColumnAttribute("");
        };

        action
            .Should().Throw<ArgumentException>()
            .WithParameterName("name");
    }
}
