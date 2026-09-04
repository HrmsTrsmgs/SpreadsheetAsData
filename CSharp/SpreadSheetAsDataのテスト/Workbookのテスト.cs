using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.Test;

public class Workbookのテスト : IDisposable
{

    readonly Workbook book1;
    readonly TemporaryExcelFiles temporaryFiles = new();

    public const string コピーパス = @"TestData\Book1-Copy.xlsx";

    public Workbookのテスト()
    {
        if (!File.Exists(コピーパス))
        {
            File.Copy(@"TestData\Book1.xlsx", コピーパス);
        }
        book1 = Workbook.Open(temporaryFiles.Copy("Book1.xlsx"));
    }

    public void Dispose()
    {
        book1.Close();
        temporaryFiles.Dispose();
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

    [Fact]
    public void ReadTableは存在しないテーブル名を指定した場合に失敗します()
    {
        using var tested = Workbook.Open(@"TestData\テーブル.xlsx");

        var action = () =>
        {
            tested.ReadTable<TestMappedRow>("missing")
                .ToArray();
        };

        action
            .Should().Throw<KeyNotFoundException>();
    }

    [Fact]
    public void Saveは開いているファイルへ変更を保存します()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A1"].Value = 9.9;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["いろいろなデータ"].Cells["A1"].Value as object)
            .Should().Be(9.9);
    }

    [Fact]
    public void SaveAsは変更したセル値を別ファイルへ保存します()
    {
        var sourcePath = temporaryFiles.Copy("Book1.xlsx");
        var savedPath = temporaryFiles.NewFilePath();

        using (var book = Workbook.Open(sourcePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A1"].Value = 9.9;
            book.SaveAs(savedPath);
        }

        using var tested = Workbook.Open(savedPath);

        (tested.Sheets["いろいろなデータ"].Cells["A1"].Value as object)
            .Should().Be(9.9);
    }

    [Fact]
    public void SaveAsは保存元ファイルを変更しません()
    {
        var sourcePath = temporaryFiles.Copy("Book1.xlsx");
        var savedPath = temporaryFiles.NewFilePath();

        using (var book = Workbook.Open(sourcePath))
        {
            book.Sheets["いろいろなデータ"].Cells["A1"].Value = 9.9;
            book.SaveAs(savedPath);
        }

        using var tested = Workbook.Open(sourcePath);

        (tested.Sheets["いろいろなデータ"].Cells["A1"].Value as object)
            .Should().Be(1.1);
    }

    [Fact]
    public void DefinedNamesはブック内の定義名を列挙します()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");

        tested.DefinedNames
            .Select(it => it.Name)
            .Should().Equal("A1", "book_cell", "book_range", "cell_name", "range_name");
    }

    [Fact]
    public void DefinedNameはブックスコープではWorksheetを返しません()
    {
        using var book = Workbook.Open(@"TestData\定義名.xlsx");

        var tested =
            (
                from definedName in book.DefinedNames
                where definedName.Name == "book_cell"
                select definedName
            ).Single();

        tested.Worksheet.Should().BeNull();
        tested.Range.Should().BeSameAs(book.Range["book_cell"]);
    }

    [Fact]
    public void DefinedNameはワークシートスコープではWorksheetを返します()
    {
        using var book = Workbook.Open(@"TestData\定義名.xlsx");
        var sheet2 = book.Sheets["Sheet2"];

        var tested =
            (
                from definedName in book.DefinedNames
                where definedName.Name == "cell_name"
                select definedName
            ).Single();

        tested.Worksheet.Should().BeSameAs(sheet2);
        tested.Range.Should().BeSameAs(sheet2.Range["cell_name"]);
    }

}
