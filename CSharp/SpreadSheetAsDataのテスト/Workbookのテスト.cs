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

    [Fact(Skip = "読み込みAPIと対になる書き込み機能をWorkbook単位で実装するときに解除する。")]
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

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableは属性で指定した列へプロパティ値を書き込みます()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.WriteTable(
                "型付き行マッピング",
                [
                    new WritableMappedRow
                    {
                        IntegerValue = 10,
                        FloatingPointValue = 1.5,
                        TextValue = "first"
                    },
                    new WritableMappedRow
                    {
                        IntegerValue = 20,
                        FloatingPointValue = 2.5,
                        TextValue = "second"
                    },
                    new WritableMappedRow
                    {
                        IntegerValue = 30,
                        FloatingPointValue = 3.5,
                        TextValue = "third"
                    }
                ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.ReadTable<TestMappedRow>("型付き行マッピング")
            .Should().BeEquivalentTo(
                [
                    new TestMappedRow
                    {
                        IntegerValue = 10,
                        FloatingPointValue = 1.5,
                        TextValue = "first"
                    },
                    new TestMappedRow
                    {
                        IntegerValue = 20,
                        FloatingPointValue = 2.5,
                        TextValue = "second"
                    },
                    new TestMappedRow
                    {
                        IntegerValue = 30,
                        FloatingPointValue = 3.5,
                        TextValue = "third"
                    }
                ],
                options => options.WithStrictOrdering());
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableは書き込み元に対応プロパティがない列を変更しません()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.WriteTable(
                "型付き行マッピング",
                [
                    new WritableIntegerOnlyRow { IntegerValue = 10 },
                    new WritableIntegerOnlyRow { IntegerValue = 20 },
                    new WritableIntegerOnlyRow { IntegerValue = 30 }
                ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);
        var rows = tested.ReadTable<TestMappedRow>("型付き行マッピング").ToArray();

        rows.Select(it => it.IntegerValue).Should().Equal(10, 20, 30);
        rows.Select(it => it.TextValue).Should().Equal("さしすせそ", "たちつてと", "なにぬねの");
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableはワークシート上の順序でデータ行を書き込みます()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.WriteTable(
                "型付き行マッピング",
                [
                    new WritableIntegerOnlyRow { IntegerValue = 10 },
                    new WritableIntegerOnlyRow { IntegerValue = 20 },
                    new WritableIntegerOnlyRow { IntegerValue = 30 }
                ]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        tested.Tables["型付き行マッピング"]
            .Rows
            .Select(it => it["数値2"].Value as object)
            .Should().Equal(10d, 20d, 30d);
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableは属性で指定した列が存在しない場合に失敗します()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            tested.WriteTable(
                "型付き行マッピング",
                [new WritableMissingColumnRow { Value = 1 }]);
        };

        action.Should().Throw<TableMappingException>();
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableは複数プロパティが同じ列を指定した場合に失敗します()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            tested.WriteTable(
                "型付き行マッピング",
                [
                    new WritableDuplicateColumnRow
                    {
                        FirstValue = 1,
                        SecondValue = 2
                    }
                ]);
        };

        action.Should().Throw<TableMappingException>();
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableは属性を付けたプロパティにpublicなgetterがない場合に失敗します()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            tested.WriteTable(
                "型付き行マッピング",
                [new WritableAttributedPropertyWithoutPublicGetterRow()]);
        };

        action.Should().Throw<TableMappingException>();
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableは属性のない書き込み専用プロパティを無視します()
    {
        var filePath = temporaryFiles.Copy("テーブル.xlsx");

        using (var book = Workbook.Open(filePath))
        {
            book.WriteTable(
                "型付き行マッピング",
                [new WritableRowWithWriteOnlyProperty { IntegerValue = 10 }]);
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Tables["型付き行マッピング"].Rows.First()["数値2"].Value as object)
            .Should().Be(10d);
    }

    [Fact(Skip = "読み込みAPIと対になる型付きテーブル書き込み機能をWorkbook単位で実装するときに解除する。")]
    public void WriteTableはセル値へ変換できないプロパティ型の場合に失敗します()
    {
        using var tested = Workbook.Open(temporaryFiles.Copy("テーブル.xlsx"));

        var action = () =>
        {
            tested.WriteTable(
                "型付き行マッピング",
                [new WritableUnsupportedValueRow { Value = DateTime.Today }]);
        };

        action.Should().Throw<TableMappingException>();
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

    public sealed class WritableMappedRow
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue { get; set; }

        [SpreadsheetColumn("数値")]
        public double FloatingPointValue { get; set; }

        [SpreadsheetColumn("文字列")]
        public string TextValue { get; set; } = "";
    }

    public sealed class WritableIntegerOnlyRow
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue { get; set; }
    }

    public sealed class WritableMissingColumnRow
    {
        [SpreadsheetColumn("missing")]
        public int Value { get; set; }
    }

    public sealed class WritableDuplicateColumnRow
    {
        [SpreadsheetColumn("数値2")]
        public int FirstValue { get; set; }

        [SpreadsheetColumn("数値2")]
        public int SecondValue { get; set; }
    }

    public sealed class WritableAttributedPropertyWithoutPublicGetterRow
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue
        {
            set { }
        }
    }

    public sealed class WritableRowWithWriteOnlyProperty
    {
        [SpreadsheetColumn("数値2")]
        public int IntegerValue { get; set; }

        public string Description
        {
            set { }
        }
    }

    public sealed class WritableUnsupportedValueRow
    {
        [SpreadsheetColumn("数値2")]
        public DateTime Value { get; set; }
    }

}
