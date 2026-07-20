using System.Reflection;
using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型推論のテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string IntegratedExcelFilePath = @"TestData\コード生成\統合.xlsx";

    [Fact]
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

    [Fact]
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

    [Fact]
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
            "生成プロパティ名とExcel列名が一致する場合に列属性を省略する仕様を実装するときに解除する。")]
    public void 列名と生成プロパティ名が一致する場合は列属性を生成しません()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        tested.Should().NotBeNull();
        tested.GetCustomAttribute<SpreadsheetColumnAttribute>()
            .Should().BeNull();
    }

    [Fact(
        Skip =
            "生成プロパティ名とExcel列名が異なる場合に既存の型付きTableマッピングへ接続する仕様を実装するときに解除する。")]
    public void 列名と生成プロパティ名が異なる場合は列属性に元のExcel列名を設定します()
    {
        var generatedProperty = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(IntegratedExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        generatedProperty.Should().NotBeNull();

        var tested =
            generatedProperty.GetCustomAttribute<SpreadsheetColumnAttribute>();

        tested.Should().NotBeNull();
        tested.Name.Should().Be("customer_id");
    }
}
