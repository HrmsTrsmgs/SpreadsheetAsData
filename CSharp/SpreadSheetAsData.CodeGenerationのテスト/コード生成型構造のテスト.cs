using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型構造のテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string IntegratedExcelFilePath = @"TestData\コード生成\統合.xlsx";

    [Theory]
    [InlineData(BasicStructureExcelFilePath, "BasicStructureBook")]
    [InlineData(IntegratedExcelFilePath, "統合Book")]
    public void 生成されたBook型はExcelファイル名に対応する型名で生成します(
        string excelFilePath,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    excelFilePath))
            .TypeNames
            .Should()
            .Contain(generatedTypeName);
    }

    [Fact]
    public void 生成されたBook型はWorkbookを継承します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GetRequiredType("BasicStructureBook")
            .Should()
            .BeAssignableTo<Workbook>();
    }

    [Theory]
    [InlineData("SalesDataSheet")]
    [InlineData("ProductMasterSheet")]
    public void 生成されたSheet型はワークシート名に対応する型名で生成します(
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .TypeNames
            .Should()
            .Contain(generatedTypeName);
    }

    [Fact]
    public void 生成されたSheet型はWorksheetを継承します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GetRequiredType("SalesDataSheet")
            .Should()
            .BeAssignableTo<Worksheet>();
    }

    [Fact]
    public void 生成されたTable型を生成します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .TypeNames
            .Should()
            .Contain("SalesDetailTable");
    }

    [Fact(
        Skip =
            "生成専用Table型が生成された行データ型を型引数にしたTableを基底型としてコンパイルできる処理を実装するときに解除する。")]
    public void 生成されたTable型は生成された行データ型を型引数とするTableを継承します()
    {
        GeneratedCodeInspection
            .SyntaxFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedType("SalesDetailTable")
            .BaseTypeName
            .Should()
            .Be("Table<SalesDetail>");
    }

    [Fact(
        Skip =
            "生成された行データ型をTableRowではなく通常のPOCOとして生成する処理を実装するときに解除する。")]
    public void 生成された行データ型はTableRowを継承しないPOCOです()
    {
        var dataType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GetRequiredType("SalesDetail");

        dataType.BaseType.Should().Be(typeof(object));
        dataType.Should().NotBeAssignableTo<TableRow>();
    }

    [Fact(
        Skip =
            "生成された行データ型の列プロパティへpublic getterとpublic setterを生成する処理を実装するときに解除する。")]
    public void 生成された行データ型の列プロパティはpublicなgetterとsetterを持ちます()
    {
        var property = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GetRequiredType("SalesDetail")
            .GetProperty("CustomerId");

        property.Should().NotBeNull();
        property!.GetMethod!.IsPublic.Should().BeTrue();
        property.SetMethod!.IsPublic.Should().BeTrue();
    }

    [Fact(
        Skip =
            "生成型をpartialとして出力し利用者のpartial定義と同時コンパイルできる処理を実装するときに解除する。")]
    public void 生成された型は別ファイルのpartial定義と共にコンパイルできます()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath);

        GeneratedSourceCompiler
            .Compile(
                [
                    .. sources,
                    """
                    namespace Generated;

                    public partial class SalesDetail
                    {
                        public bool AddedByUser => true;
                    }
                    """
                ])
            .GetRequiredType("SalesDetail")
            .GetProperty("AddedByUser")
            .Should()
            .NotBeNull();
    }
}
