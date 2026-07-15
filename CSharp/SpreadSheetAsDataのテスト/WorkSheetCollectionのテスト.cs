using System.Collections;
using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class WorkSheetCollectionのテスト
{
    readonly WorksheetCollection tested;

    public WorkSheetCollectionのテスト()
    {
        using var book = Workbook.Open(@"TestData\Book1.xlsx");
        tested = book.Sheets;
    }

    [Fact]
    public void インデクサに数字でアクセスできます()
    {
        tested[0].Name.Should().Be("Sheet1");
    }

    [Fact]
    public void インデクサにシート名でアクセスできます()
    {
        tested["Sheet1"].Should().BeSameAs(tested[0]);
    }

    [Fact]
    public void Countで個数を取得できます()
    {
        tested.Count.Should().Be(3);
    }

    [Fact]
    public void Keysでシート名の一覧が取得できます()
    {
        tested.Keys.Should().Contain("Sheet1");
        tested.Keys.Should().Contain("Sheet2");
    }

    [Fact]
    public void Valuesでシートの一覧が取得できます()
    {
        tested.Values.Should().Contain(tested[0]);
        tested.Values.Should().Contain(tested[1]);
    }

    [Fact]
    public void ContainsKeyでシート名の有無が確認できます()
    {
        tested.ContainsKey("Sheet1").Should().BeTrue();
        tested.ContainsKey("").Should().BeFalse();
    }

    [Fact]
    public void TryGetValueでシート名の有無が確認しつつシートの取得ができます()
    {
        tested.TryGetValue("Sheet1", out var sheet).Should().BeTrue();
        sheet.Should().BeSameAs(tested[0]);
        tested.TryGetValue("", out _).Should().BeFalse();
    }

    [Fact]
    public void シートはforeachで取得できます()
    {
        var index = 0;
        foreach (var sheet in tested)
        {
            sheet.Should().Be(tested[index++]);
        }
    }

    [Fact]
    public void 非ジェネリックのforeachでがシートが取得できます()
    {
        var index = 0;
        foreach (var sheet in (IEnumerable)tested)
        {
            sheet.Should().Be(tested[index++]);
        }
    }

    [Fact]
    public void Dictionaryに対するのforeachでがシートが取得できます()
    {
        var index = 0;
        foreach (var entry in (IReadOnlyDictionary<string, Worksheet>)tested)
        {
            entry.Key.Should().Be(tested[index].Name);
            entry.Value.Should().Be(tested[index++]);
        }
    }
}
