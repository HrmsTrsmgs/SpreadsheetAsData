using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class SpreadsheetColumnAttributeのテスト
{
    [Fact(Skip = "SpreadsheetColumnAttributeが指定された列名を保持する実装時に解除する。")]
    public void Nameはコンストラクターで指定した列名を返します()
    {
        new SpreadsheetColumnAttribute("int")
            .Name
            .Should()
            .Be("int");
    }

    [Fact(Skip = "SpreadsheetColumnAttributeの列名にnullを指定できないようにするときに解除する。")]
    public void コンストラクターはnullの列名を拒否します()
    {
        var action = () =>
        {
            _ = new SpreadsheetColumnAttribute(null!);
        };

        action
            .Should()
            .Throw<ArgumentNullException>()
            .WithParameterName("name");
    }

    [Fact(Skip = "SpreadsheetColumnAttributeの列名に空文字列を指定できないようにするときに解除する。")]
    public void コンストラクターは空の列名を拒否します()
    {
        var action = () =>
        {
            _ = new SpreadsheetColumnAttribute("");
        };

        action
            .Should()
            .Throw<ArgumentException>()
            .WithParameterName("name");
    }
}
