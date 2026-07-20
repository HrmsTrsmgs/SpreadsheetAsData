using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成アクセスのテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";

    [Fact]
    public void Bookは各ワークシートを型付きプロパティとして公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var tested = generatedAssembly
            .GeneratedType("BasicStructureBook")
            .GetProperty("SalesData");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            generatedAssembly.GeneratedType("SalesDataSheet"));
    }

    [Fact]
    public void Bookは各Excelテーブルを型付きプロパティとして公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var tested = generatedAssembly
            .GeneratedType("BasicStructureBook")
            .GetProperty("SalesDetail");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            generatedAssembly.GeneratedType("SalesDetailTable"));
    }

    [Fact(
        Skip =
            "生成BookをWorkbookとして扱える継承構造と既存非型付きAPIの利用を実装するときに解除する。")]
    public void 生成されたBook型からWorkbookの非型付きAPIも使用できます()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Workbook>("BasicStructureBook");

        tested.Sheets.Should().NotBeNull();
        tested.Tables.Should().NotBeNull();
        tested.Cell.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested["SalesData"].Should().NotBeNull();
    }

    [Fact(
        Skip =
            "Sheetからそのシートに属するExcelテーブルだけを型付きプロパティとして取得する生成処理を実装するときに解除する。")]
    public void Sheetはそのシートに属するExcelテーブルを型付きプロパティとして公開します()
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var sheetType = generatedAssembly.GeneratedType("SalesDataSheet");
        var tested = sheetType.GetProperty("SalesDetail");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            generatedAssembly.GeneratedType("SalesDetailTable"));
        sheetType.GetProperty("ProductList").Should().BeNull();
    }

    [Fact(
        Skip =
            "生成SheetをWorksheetとして扱える継承構造と既存非型付きAPIの利用を実装するときに解除する。")]
    public void 生成されたSheet型からWorksheetの非型付きAPIも使用できます()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedInstance<Worksheet>("SalesDataSheet");

        tested.Name.Should().NotBeNull();
        tested.Book.Should().NotBeNull();
        tested.Cell.Should().NotBeNull();
        tested.Range.Should().NotBeNull();
        tested.Cells.Should().NotBeNull();
    }
}
