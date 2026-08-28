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
    public void ブックスコープの単一セル定義名をBookのCellプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("定義名Book")
            .GetProperty("MainCell");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(Cell));
    }

    [Fact]
    public void ブックスコープの複数セル定義名をBookのCellRangeプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("定義名Book")
            .GetProperty("MainRange");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(CellRange));
    }

    [Fact]
    public void シートローカルの単一セル定義名をSheetのCellプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("SalesDataSheet")
            .GetProperty("LocalCell");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(Cell));
    }

    [Fact]
    public void シートローカルの複数セル定義名をSheetのCellRangeプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath))
            .GeneratedType("SalesDataSheet")
            .GetProperty("LocalRange");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(CellRange));
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

    [Fact(Skip = "読み込みコード生成と対になる定義名書き込み機能を実装するときに解除する。")]
    public void 生成されたBook型のCellプロパティから値を書き込めます()
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

            bookAccessor.MainCell.Value = "generated";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Cell["book_cell"].Value as object)
            .Should().Be("generated");
    }

    [Fact(Skip = "読み込みコード生成と対になる定義名書き込み機能を実装するときに解除する。")]
    public void 生成されたSheet型のCellプロパティから値を書き込めます()
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

            bookAccessor.SalesData.LocalCell.Value = "generated";
            book.Save();
        }

        using var tested = Workbook.Open(filePath);

        (tested.Sheets["SalesData"].Cell["local_cell"].Value as object)
            .Should().Be("generated");
    }
}
