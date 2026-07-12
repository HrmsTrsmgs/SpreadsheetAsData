using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Workbookのテスト : IDisposable
{

    readonly Workbook book1;

    public const string コピーパス = @"TestData\Book1-Copy.xlsx";

    public Workbookのテスト()
    {
        if (!File.Exists(コピーパス))
        {
            File.Copy(@"TestData\Book1.xlsx", コピーパス);
        }
        book1 = Workbook.Open(@"TestData\Book1.xlsx");
    }

    public void Dispose()
    {
        book1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Openはファイルを束縛します()
    {
        var tested = Workbook.Open(コピーパス);
        try
        {
            FluentActions.Invoking(
                () => File.Delete(コピーパス)
            ).Should().Throw<IOException>();
        }
        finally
        {
            tested.Close();
        }
    }

    [Fact]
    public void Closeはファイルの束縛を解除します()
    {
        var tested = Workbook.Open(コピーパス);

        tested.Close();
        FluentActions.Invoking(
            () => File.Delete(コピーパス)
        ).Should().NotThrow();
    }

    [Fact]
    public void Disposeはファイルの束縛を解除します()
    {
        using (var tested = Workbook.Open(コピーパス))
        {
        }
        FluentActions.Invoking(
            () => File.Delete(コピーパス)
        ).Should().NotThrow();
    }

    [Fact]
    public void Sheetsでシートが取得できます()
    {
        book1.Sheets.Count.Should().Be(3);
    }

    [Fact]
    public void Sheetsに数字を指定してシートが取得できます()
    {
        book1.Sheets[0].Name.Should().Be("Sheet1");
    }

    [Fact]
    public void Sheetsにシート名を指定してシートが取得できます()
    {
        book1.Sheets["Sheet1"].Name.Should().Be("Sheet1");
    }

    [Fact]
    public void インデクサに数字を指定してシートが取得できます()
    {
        book1[0].Name.Should().Be("Sheet1");
    }

    [Fact]
    public void インデクサにシート名を指定してシートが取得できます()
    {
        book1["Sheet1"].Name.Should().Be("Sheet1");
    }

}
