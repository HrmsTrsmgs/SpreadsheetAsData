using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Marimo.SpreadSheetAsData;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成定義名のテスト
{
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

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

    [Fact(
        Skip =
            "シートローカルの複数セル定義名をSheetのCellRange取得プロパティとして生成するときに解除する。")]
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

    [Fact(
        Skip =
            "同じ定義名をブックスコープとシートローカルで別々のプロパティとして生成するときに解除する。")]
    public void ブックスコープとシートローカルで同じ定義名を区別して生成します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(DefinedNamesExcelFilePath));

        generatedAssembly.GeneratedType("定義名Book")
            .GetProperty("Total").Should().NotBeNull();
        generatedAssembly.GeneratedType("SalesDataSheet")
            .GetProperty("Total").Should().NotBeNull();
    }
}
