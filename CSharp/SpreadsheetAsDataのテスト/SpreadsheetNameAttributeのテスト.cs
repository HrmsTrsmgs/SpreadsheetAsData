using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public class SpreadsheetNameAttributeのテスト
{
    [Fact]
    public void Nameはコンストラクターで指定したExcel上の名前を返します()
    {
        new Marimo.SpreadsheetAsData.SpreadsheetNameAttribute("name")
            .Name
            .Should().Be("name");
    }

    [Fact]
    public void コンストラクターはnullの名前を拒否します()
    {
        var action = () =>
        {
            // nullableを認識しない呼び出し元からのnullを、実行時にも拒否することを検証します。
            _ = new SpreadsheetNameAttribute(null!);
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
            _ = new SpreadsheetNameAttribute("");
        };

        action
            .Should().Throw<ArgumentException>()
            .WithParameterName("name");
    }
}
