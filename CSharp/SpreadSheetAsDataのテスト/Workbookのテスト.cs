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
    public void OpenはMemoryStream上のブックを開きます()
    {
        using var stream = new MemoryStream(
            File.ReadAllBytes(@"TestData\Book1.xlsx"));
        using var tested = Workbook.Open(stream);

        tested.Sheets.Keys
            .Should().Equal("Sheet1", "Sheet2", "いろいろなデータ");
    }

    [Fact]
    public void Openは再圧縮でサイズが増えるブックも固定容量のMemoryStreamで開いて閉じられます()
    {
        using var stream = new MemoryStream(
            File.ReadAllBytes(@"TestData\再圧縮でサイズが増えるブック.xlsx"));

        var tested = () =>
        {
            using var book = Workbook.Open(stream);
        };

        tested.Should().NotThrow();
    }

    [Fact]
    public void OpenはFileStream上のブックを開きます()
    {
        using var stream = File.Open(
            temporaryFiles.Copy("Book1.xlsx"),
            FileMode.Open,
            FileAccess.ReadWrite);
        using var tested = Workbook.Open(stream);

        tested.Sheets.Keys
            .Should().Equal("Sheet1", "Sheet2", "いろいろなデータ");
    }

    [Fact]
    public void Openは検証指定を省略した場合不正なOOXMLも開きます()
    {
        var action = () =>
        {
            using var book = Workbook.Open(
                temporaryFiles.Copy("不正なOOXML.xlsx"));
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Openは検証しない場合不正なOOXMLも開きます()
    {
        var action = () =>
        {
            using var book = Workbook.Open(
                temporaryFiles.Copy("不正なOOXML.xlsx"),
                validate: false);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Openは検証する場合正常なOOXMLを開きます()
    {
        var action = () =>
        {
            using var book = Workbook.Open(
                temporaryFiles.Copy("文字列セル.xlsx"),
                validate: true);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Openは検証する場合不正なOOXMLで失敗します()
    {
        var action = () =>
        {
            using var book = Workbook.Open(
                temporaryFiles.Copy("不正なOOXML.xlsx"),
                validate: true);
        };

        action.Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void Disposeは呼び出し側から渡されたStreamを閉じません()
    {
        using var stream = new MemoryStream(
            File.ReadAllBytes(@"TestData\Book1.xlsx"));
        var tested = Workbook.Open(stream);

        tested.Dispose();

        stream.CanRead.Should().BeTrue();
        stream.CanWrite.Should().BeTrue();
    }

    [Fact]
    public void Stream版でも定義名からオブジェクトを読み込みます()
    {
        using var stream = new MemoryStream(
            File.ReadAllBytes(@"TestData\定義名.xlsx"));
        using var tested = Workbook.Open(stream);

        tested.Read<WorkbookData>()
            .Should().BeEquivalentTo(
                new WorkbookData
                {
                    CustomerName = "山田太郎"
                });
    }

    [Fact]
    public void Stream版のSaveAsは元のStreamを変更せず別ファイルへ保存します()
    {
        var original = File.ReadAllBytes(@"TestData\定義名.xlsx");
        using var stream = new MemoryStream(original.ToArray());
        var savedPath = temporaryFiles.NewFilePath();

        using (var book = Workbook.Open(stream))
        {
            book.Replace(
                new WorkbookData
                {
                    CustomerName = "佐藤花子"
                });
            book.SaveAs(savedPath);
            stream.ToArray().Should().Equal(original);
        }

        stream.ToArray().Should().Equal(original);
        using var tested = Workbook.Open(savedPath);

        (tested.Cell["CustomerName"].Value as object)
            .Should().Be("佐藤花子");
    }

    [Fact]
    public void DisposeはSaveしていない変更を元のStreamへ書き込みません()
    {
        var original = File.ReadAllBytes(@"TestData\定義名.xlsx");
        using var stream = new MemoryStream(original.ToArray());

        using (var book = Workbook.Open(stream))
        {
            book.Cell["CustomerName"].Value = "保存しない変更";
        }

        stream.ToArray().Should().Equal(original);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Stream版のSaveは拡張可否によらず元のStreamを変更せずに拒否します(bool expandable)
    {
        var original = File.ReadAllBytes(@"TestData\定義名.xlsx");
        using var stream = expandable
            ? new MemoryStream()
            : new MemoryStream(new byte[original.Length]);
        stream.Write(original);
        stream.Position = 0;

        using (var book = Workbook.Open(stream))
        {
            book.Cell["CustomerName"].Value = "佐藤花子";
            var tested = () => book.Save();

            tested.Should().Throw<NotSupportedException>();
            stream.ToArray().Should().Equal(original);
        }

        stream.ToArray().Should().Equal(original);
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
    public void Saveはブックを閉じる前に元ファイルへ変更を反映します()
    {
        var filePath = temporaryFiles.Copy("Book1.xlsx");
        using var book = Workbook.Open(filePath);
        book.Sheets["いろいろなデータ"].Cells["A1"].Value = 9.9;
        book.Save();

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
    public void Readはプロパティ名と同じブックスコープの単一セル定義名からオブジェクトを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");

        tested.Read<WorkbookData>()
            .Should().BeEquivalentTo(
                new WorkbookData
                {
                    CustomerName = "山田太郎"
                });
    }

    [Fact]
    public void Readは空白の単一セル定義名をstringプロパティの空文字列として読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");
        tested.Cell["CustomerName"].Value = null;

        tested.Read<WorkbookData>().CustomerName.Should().BeEmpty();
    }

    [Fact]
    public void ReadはSpreadSheetName属性で指定した定義名からオブジェクトを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");

        tested.Read<AttributedWorkbookData>()
            .Should().BeEquivalentTo(
                new AttributedWorkbookData
                {
                    Name = "山田太郎"
                });
    }

    [Fact]
    public void ReadはSpreadSheetName属性で指定したシートローカルの単一セル定義名からオブジェクトを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");
        tested.Sheets["Sheet2"].Cell["cell_name"].Value = "シートローカル";

        tested.Read<SheetScopedWorkbookData>()
            .Value.Should().Be("シートローカル");
    }

    [Fact]
    public void Readは自動変換したプロパティ名と一致するシートローカル定義名からオブジェクトを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");
        tested.Sheets["Sheet2"].Cell["cell_name"].Value = "シートローカル";

        tested.Read<ConventionWorkbookData>()
            .CellName.Should().Be("シートローカル");
    }

    [Fact]
    public void ReadはSpreadSheetName属性で指定した複数セル定義名からオブジェクトを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");
        var expected = tested.Range["book_range"].Values
            .Select(it => it.ToArray())
            .ToArray();

        tested.Read<AttributedWorkbookRangeData>()
            .Values
            .Select(it => it.ToArray())
            .Should().BeEquivalentTo(
                expected,
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void Readはプロパティ名と同じExcelテーブルから行データを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\テーブル.xlsx");

        tested.Read<WorkbookTableData>()
            .型付き行マッピング
            .Select(it => it.IntegerValue)
            .Should().Equal(1, 2, 3);
    }

    [Fact]
    public void Readはテーブルのnullableプロパティへ値と空白を読み込みます()
    {
        using var book = Workbook.Open(@"TestData\テーブル.xlsx");
        var tested = book.Read<NullableWorkbookTableData>().空白数値マッピング.ToArray();

        tested.Select(it => it.数値).Should().Equal(0, null);
        tested.Select(it => it.小数).Should().Equal(1.5, null);
        tested.Select(it => it.真偽値).Should().Equal(false, null);
    }

    [Fact]
    public void Replaceはテーブルのnullableプロパティの値とnullを書き込みます()
    {
        using var book = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        book.Replace(
            new NullableWorkbookTableData
            {
                空白数値マッピング =
                [
                    new NullableTableRowData { 数値 = null, 小数 = null, 真偽値 = null },
                    new NullableTableRowData { 数値 = 10, 小数 = 2.5, 真偽値 = true }
                ]
            });

        var tested = book.Tables["空白数値マッピング"].Rows.ToArray();
        tested.Select(it => it["数値"].Value as object)
            .Should().Equal(new BlankValue(), 10d);
        tested.Select(it => it["小数"].Value as object)
            .Should().Equal(new BlankValue(), 2.5);
        tested.Select(it => it["真偽値"].Value as object)
            .Should().Equal(new BlankValue(), true);
    }

    [Fact]
    public void ReadはSpreadSheetName属性で指定したExcelテーブルから行データを読み込みます()
    {
        using var tested = Workbook.Open(@"TestData\テーブル.xlsx");

        tested.Read<AttributedWorkbookTableData>()
            .OrderLines
            .Select(it => it.IntegerValue)
            .Should().Equal(1, 2, 3);
    }

    [Fact]
    public void Replaceはプロパティ名と同じExcelテーブルへ行データを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        tested.Replace(
            new WorkbookTableData
            {
                型付き行マッピング =
                [
                    new TestMappedRow { IntegerValue = 4 },
                    new TestMappedRow { IntegerValue = 5 },
                    new TestMappedRow { IntegerValue = 6 }
                ]
            });

        tested.Tables["型付き行マッピング"].Rows
            .Select(it => it["数値2"].Value)
            .Should().Equal(4d, 5d, 6d);
    }

    [Fact]
    public void ReplaceはSpreadSheetName属性で指定したExcelテーブルへ行データを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        tested.Replace(
            new AttributedWorkbookTableData
            {
                OrderLines =
                [
                    new TestMappedRow { IntegerValue = 4 },
                    new TestMappedRow { IntegerValue = 5 },
                    new TestMappedRow { IntegerValue = 6 }
                ]
            });

        tested.Tables["型付き行マッピング"].Rows
            .Select(it => it["数値2"].Value)
            .Should().Equal(4d, 5d, 6d);
    }

    [Fact]
    public void Replaceはプロパティ名と同じブックスコープの単一セル定義名へオブジェクトを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("定義名.xlsx"));

        tested.Replace(
            new WorkbookData
            {
                CustomerName = "佐藤花子"
            });

        (tested.Cell["CustomerName"].Value as object).Should().Be("佐藤花子");
    }

    [Fact]
    public void ReplaceはSpreadSheetName属性で指定した定義名へオブジェクトを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("定義名.xlsx"));

        tested.Replace(
            new AttributedWorkbookData
            {
                Name = "佐藤花子"
            });

        (tested.Cell["CustomerName"].Value as object).Should().Be("佐藤花子");
    }

    [Fact]
    public void ReplaceはSpreadSheetName属性で指定したシートローカルの単一セル定義名へオブジェクトを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("定義名.xlsx"));

        tested.Replace(
            new SheetScopedWorkbookData
            {
                Value = "シートローカル"
            });

        (tested.Sheets["Sheet2"].Cell["cell_name"].Value as object)
            .Should().Be("シートローカル");
    }

    [Fact]
    public void Replaceは自動変換したプロパティ名と一致するシートローカル定義名へオブジェクトを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("定義名.xlsx"));

        tested.Replace(
            new ConventionWorkbookData
            {
                CellName = "シートローカル"
            });

        (tested.Sheets["Sheet2"].Cell["cell_name"].Value as object)
            .Should().Be("シートローカル");
    }

    [Fact]
    public void ReplaceはSpreadSheetName属性で指定した複数セル定義名へオブジェクトを書き込みます()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("定義名.xlsx"));
        IEnumerable<IEnumerable<object?>> replacement =
        [
            ["a", "b"],
            ["c", "d"],
            ["e", "f"],
            ["g", "h"],
            ["i", "j"]
        ];

        tested.Replace(
            new AttributedWorkbookRangeData
            {
                Values = replacement
            });

        tested.Range["book_range"].Values
            .Should().BeEquivalentTo(
                replacement,
                options => options.WithStrictOrdering());
    }

    [Fact]
    public void DefinedNamesはブック内の定義名を列挙します()
    {
        using var tested = Workbook.Open(@"TestData\定義名.xlsx");

        tested.DefinedNames
            .Select(it => it.Name)
            .Should().Equal(
                "A1",
                "book_cell",
                "book_range",
                "cell_name",
                "range_name",
                "CustomerName");
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

    public sealed class WorkbookData
    {
        public string CustomerName { get; set; } = "";
    }

    public sealed class AttributedWorkbookData
    {
        [SpreadSheetName("CustomerName")]
        public string Name { get; set; } = "";
    }

    public sealed class AttributedWorkbookRangeData
    {
        [SpreadSheetName("book_range")]
        public IEnumerable<IEnumerable<object?>> Values { get; set; } = [];
    }

    public sealed class SheetScopedWorkbookData
    {
        [SpreadSheetName("cell_name", WorksheetName = "Sheet2")]
        public string Value { get; set; } = "";
    }

    public sealed class ConventionWorkbookData
    {
        public string CellName { get; set; } = "";
    }

    public sealed class WorkbookTableData
    {
        public IEnumerable<TestMappedRow> 型付き行マッピング { get; set; } = [];
    }

    public sealed class NullableWorkbookTableData
    {
        public IEnumerable<NullableTableRowData> 空白数値マッピング { get; set; } = [];
    }

    public sealed class NullableTableRowData
    {
        public int? 数値 { get; set; }
        public double? 小数 { get; set; }
        public bool? 真偽値 { get; set; }
    }

    public sealed class AttributedWorkbookTableData
    {
        [SpreadSheetName("型付き行マッピング")]
        public IEnumerable<TestMappedRow> OrderLines { get; set; } = [];
    }

}
