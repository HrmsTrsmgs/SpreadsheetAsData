using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class DefinedNameのテスト
{
    [Fact]
    public void ToStringはExcel上の定義名を返します()
    {
        using var book = Workbook.Open(@"TestData\定義名.xlsx");

        book.DefinedNames
            .Single(it => it.Name == "book_cell")
            .ToString().Should().Be("book_cell");
        book.DefinedNames
            .Single(it => it.Name == "cell_name")
            .ToString().Should().Be("cell_name");
    }
}
