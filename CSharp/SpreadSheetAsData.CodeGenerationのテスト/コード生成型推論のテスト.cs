using System.Reflection;
using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型推論のテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";

    [Fact(
        Skip =
            "整数値だけを持つExcel列をintプロパティとして生成するときに解除する。")]
    public void 整数値だけを持つ列をintプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(int));
    }

    [Fact(
        Skip =
            "小数値を持つExcel列をdoubleプロパティとして生成するときに解除する。")]
    public void 小数値を持つ列をdoubleプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("Amount");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(double));
    }

    [Fact(
        Skip =
            "文字列値を持つExcel列をstringプロパティとして生成するときに解除する。")]
    public void 文字列値を持つ列をstringプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("Description");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(string));
    }

    [Fact(
        Skip =
            "生成列プロパティへ元のExcel列名をSpreadsheetColumn属性として出力するときに解除する。")]
    public void 生成された列プロパティに元のExcel列名を設定します()
    {
        var generatedProperty = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        generatedProperty.Should().NotBeNull();

        var tested =
            generatedProperty.GetCustomAttribute<SpreadsheetColumnAttribute>();

        tested.Should().NotBeNull();
        tested.Name.Should().Be("customer_id");
    }
}
