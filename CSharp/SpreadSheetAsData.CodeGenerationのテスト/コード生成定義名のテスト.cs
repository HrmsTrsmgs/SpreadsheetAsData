using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成定義名のテスト : IDisposable
{
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

    readonly TemporaryExcelFiles temporaryFiles = new();

    public void Dispose()
    {
        temporaryFiles.Dispose();
        GC.SuppressFinalize(this);
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
        tested.PropertyType.Should().Be(typeof(object));
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
        tested.PropertyType.Should().Be(typeof(IEnumerable<IEnumerable<object?>>));
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
        tested.PropertyType.Should().Be(typeof(object));
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
        tested.PropertyType.Should().Be(typeof(IEnumerable<IEnumerable<object?>>));
    }

    [Fact]
    public void ブックスコープとシートローカルで同じ定義名を区別して生成します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath));

        generatedAssembly.GeneratedType("定義名Book")
            .GetProperty("Total").Should().NotBeNull();
        generatedAssembly.GeneratedType("SalesDataSheet")
            .GetProperty("Total").Should().NotBeNull();
    }

    [Fact]
    public void 生成されたBook型の単一セル定義名プロパティへ値を直接書き込めます()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);

        using (var book = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = book;

            bookAccessor.MainCell = "generated";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Cell["main_cell"].Value as object)
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

        using (var book = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = book;

            bookAccessor.MainRange = replacement;
            book.Save();
        }

        using var tested = Workbook.Open(filePath);
        var rows = tested.Range["main_range"].Values
            .Select(it => it.ToArray())
            .ToArray();

        rows[0].Should().Equal("changed", "values");
        rows[1].Should().Equal(300d, 400d);
    }

    [Fact]
    public void 生成されたSheet型の単一セル定義名プロパティへ値を直接書き込めます()
    {
        var filePath = temporaryFiles.Copy(DefinedNamesExcelFilePath);

        using (var book = GeneratedCodeInspection
                   .AssemblyFrom(
                       GeneratedCodeInspection.GenerateSources(
                           DefinedNamesExcelFilePath))
                   .GeneratedInstance<Workbook>(
                       "定義名Book",
                       filePath))
        {
            dynamic bookAccessor = book;

            bookAccessor.SalesData.LocalCell = "generated";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["sales_data"].Cell["local_cell"].Value as object)
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
}
