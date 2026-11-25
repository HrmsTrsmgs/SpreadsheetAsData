using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Marimo.SpreadsheetAsData;
using Marimo.SpreadsheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadsheetAsData.CodeGeneration.Test;

public sealed class コード生成定義名のテスト : IDisposable
{
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\衝突なし\定義名.xlsx";
    const string SameNameInDifferentScopesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

    readonly TemporaryExcelFiles temporaryFiles = new();

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void Dataの定義名はブックスコープの後に全シートの単一セルと範囲の順で生成します()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);
        using (var document = SpreadsheetDocument.Open(filePath, true))
        {
            var definedNames = document.WorkbookPart?.Workbook?.DefinedNames
                ?? throw new InvalidOperationException("テストブックに定義名がありません。");
            definedNames.AppendChild(new Spreadsheet.DefinedName
            {
                Name = "product_cell",
                LocalSheetId = 1U,
                Text = "product_master!$A$2"
            });
        }

        GeneratedCodeInspection.SyntaxFrom(GeneratedCodeInspection.GenerateSources(filePath))
            .GeneratedType($"{Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier()}Data")
            .PropertyNames.Should().Equal(
                "MainCell", "Total", "MainRange",
                "SheetTotal", "LocalCell", "ProductCell", "LocalRange",
                "SalesDetail", "ProductList");
    }

    [Fact]
    public void ブックスコープの単一セル定義名をBookの値プロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("定義名Book")
            .GetProperty("MainCell");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be<object>();
    }

    [Fact]
    public void ブックスコープの複数セル定義名をBookの値プロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("定義名Book")
            .GetProperty("MainRange");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be<IEnumerable<IEnumerable<object?>>>();
    }

    [Fact]
    public void シートローカルの単一セル定義名をSheetの値プロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("SalesDataSheet")
            .GetProperty("LocalCell");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be<object>();
    }

    [Fact]
    public void シートローカルの複数セル定義名をSheetの値プロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("SalesDataSheet")
            .GetProperty("LocalRange");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be<IEnumerable<IEnumerable<object?>>>();
    }

    [Fact]
    public void ブックスコープとシートローカルで同じ定義名を区別して生成します()
    {
        var generatedSources = GeneratedCodeInspection.GenerateSources(
            SameNameInDifferentScopesExcelFilePath);

        generatedSources.TypeDeclaration("定義名Book")
            .PropertyDeclaration("Total").Should().NotBeNull();
        generatedSources.TypeDeclaration("SalesDataSheet")
            .PropertyDeclaration("Total").Should().NotBeNull();
    }

    [Fact]
    public void 生成されたBook型の単一セル定義名プロパティへ値を直接書き込めます()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);

        using (var tested = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = tested;

            bookAccessor.MainCell = "generated";
            tested.Save();
        }

        using var savedBook = Workbook.Open(filePath);

        (savedBook.Cell["main_cell"].Value as object)
            .Should().Be("generated");
    }

    [Fact]
    public void 生成されたBook型の単一セル定義名プロパティから値を直接読み取れます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesExcelFilePath);
        dynamic bookAccessor = book;

        object? tested = bookAccessor.MainCell;

        tested.Should().Be("main");
    }

    [Fact]
    public void 生成されたBook型の複数セル定義名プロパティから値を直接読み取れます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesExcelFilePath);
        dynamic bookAccessor = book;

        IEnumerable<IEnumerable<object?>> tested = bookAccessor.MainRange;
        var rows = tested
            .Select(it => it.ToArray())
            .ToArray();

        rows.Should().HaveCount(2);
        rows[0].Should().Equal("main", "range");
        rows[1].Should().Equal(100d, 200d);
    }

    [Fact]
    public void 生成されたBook型の複数セル定義名プロパティへ値を直接書き込めます()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);
        IEnumerable<IEnumerable<object?>> replacement =
        [
            ["changed", "values"],
            [300, 400]
        ];

        using (var tested = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = tested;

            bookAccessor.MainRange = replacement;
            tested.Save();
        }

        using var savedBook = Workbook.Open(filePath);
        var rows = savedBook.Range["main_range"].Values
            .Select(it => it.ToArray())
            .ToArray();

        rows[0].Should().Equal("changed", "values");
        rows[1].Should().Equal(300d, 400d);
    }

    [Fact]
    public void 生成されたSheet型の単一セル定義名プロパティへ値を直接書き込めます()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);

        using (var tested = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = tested;

            bookAccessor.SalesData.LocalCell = "generated";
            tested.Save();
        }

        using var savedBook = Workbook.Open(filePath);

        (savedBook.Sheets["sales_data"].Cell["local_cell"].Value as object)
            .Should().Be("generated");
    }

    [Fact]
    public void 生成されたSheet型の単一セル定義名プロパティから値を直接読み取れます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesExcelFilePath);
        dynamic bookAccessor = book;

        object? tested = bookAccessor.SalesData.LocalCell;

        tested.Should().Be(1d);
    }

    [Fact]
    public void 生成されたSheet型の複数セル定義名プロパティから値を直接読み取れます()
    {
        using var book = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedInstance<Workbook>(
                "定義名Book",
                DefinedNamesExcelFilePath);
        dynamic bookAccessor = book;

        IEnumerable<IEnumerable<object?>> tested = bookAccessor.SalesData.LocalRange;
        var rows = tested
            .Select(it => it.ToArray())
            .ToArray();

        rows.Should().HaveCount(2);
        rows[0].Should().Equal(1d, 10.5d);
        rows[1].Should().Equal(2d, 20.5d);
    }

    [Fact]
    public void 生成されたSheet型の複数セル定義名プロパティへ値を直接書き込めます()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);
        IEnumerable<IEnumerable<object?>> replacement =
        [
            [3, 30.5],
            [4, 40.5]
        ];

        using (var tested = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = tested;

            bookAccessor.SalesData.LocalRange = replacement;
            tested.Save();
        }

        using var savedBook = Workbook.Open(filePath);
        var rows = savedBook.Sheets["sales_data"].Range["local_range"].Values
            .Select(it => it.ToArray())
            .ToArray();

        rows[0].Should().Equal(3d, 30.5d);
        rows[1].Should().Equal(4d, 40.5d);
    }
}
