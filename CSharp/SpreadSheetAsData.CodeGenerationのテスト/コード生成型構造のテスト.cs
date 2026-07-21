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
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(excelFilePath))
            .DefinedTypes
            .Select(it => it.Name)
            .Should().Contain(generatedTypeName);
    }

    [Fact]
    public void 生成されたBook型はWorkbookを継承します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedType("BasicStructureBook")
            .Should().BeAssignableTo<Workbook>();
    }

    [Theory]
    [InlineData("SalesDataSheet")]
    [InlineData("ProductMasterSheet")]
    public void 生成されたSheet型はワークシート名に対応する型名で生成します(
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .DefinedTypes
            .Select(it => it.Name)
            .Should().Contain(generatedTypeName);
    }

    [Fact]
    public void 生成されたSheet型はWorksheetを継承します()
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedType("SalesDataSheet")
            .Should().BeAssignableTo<Worksheet>();
    }

    [Theory]
    [InlineData("SalesDetail", "SalesDetailTable")]
    [InlineData("ProductList", "ProductListTable")]
    public void 生成されたTable型をExcelテーブル名に対応する型名で生成します(
        string excelTableName,
        string generatedTypeName)
    {
        GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .DefinedTypes
            .Select(it => it.Name)
            .Should().Contain(
                generatedTypeName,
                "Excelテーブル {0} から生成される型名だから",
                excelTableName);
    }

    [Theory]
    [InlineData("SalesDetailTable", "SalesDetail")]
    [InlineData("ProductListTable", "ProductList")]
    public void 生成されたTable型は生成された行データ型を型引数とするTableを継承します(
        string generatedTypeName,
        string rowTypeName)
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));

        var tested = generatedAssembly.GeneratedType(generatedTypeName).BaseType;

        tested.Should().NotBeNull();
        tested.GetGenericTypeDefinition().Should().Be(typeof(Table<>));
        tested.GenericTypeArguments.Should().Equal(
            generatedAssembly.GeneratedType(rowTypeName));
    }

    [Fact]
    public void 生成された行データ型はTableRowを継承しないPOCOです()
    {
        var dataType = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail");

        dataType.BaseType.Should().Be(typeof(object));
        dataType.Should().NotBeAssignableTo<TableRow>();
    }

    [Fact]
    public void 生成された行データ型の列プロパティはpublicなgetterとsetterを持ちます()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        tested.Should().NotBeNull();

        tested.GetMethod.Should().NotBeNull();
        tested.GetMethod.IsPublic.Should().BeTrue();

        tested.SetMethod.Should().NotBeNull();
        tested.SetMethod.IsPublic.Should().BeTrue();
    }

    [Fact]
    public void 生成されたBook型は別ファイルのpartial定義と共にコンパイルできます()
    {
        GeneratedSourceCompiler
            .Compile(
                [
                    .. GeneratedCodeInspection.GenerateSources(
                        BasicStructureExcelFilePath),
                    """
                    namespace Generated;

                    public partial class BasicStructureBook
                    {
                        public bool AddedByUser => true;
                    }
                    """
                ])
            .GeneratedType("BasicStructureBook")
            .GetProperty("AddedByUser")
            .Should().NotBeNull();
    }

    [Fact]
    public void 生成されたSheet型は別ファイルのpartial定義と共にコンパイルできます()
    {
        GeneratedSourceCompiler
            .Compile(
                [
                    .. GeneratedCodeInspection.GenerateSources(
                        BasicStructureExcelFilePath),
                    """
                    namespace Generated;

                    public partial class SalesDataSheet
                    {
                        public bool AddedByUser => true;
                    }
                    """
                ])
            .GeneratedType("SalesDataSheet")
            .GetProperty("AddedByUser")
            .Should().NotBeNull();
    }

    [Fact]
    public void 生成されたTable型は別ファイルのpartial定義と共にコンパイルできます()
    {
        GeneratedSourceCompiler
            .Compile(
                [
                    .. GeneratedCodeInspection.GenerateSources(
                        BasicStructureExcelFilePath),
                    """
                    namespace Generated;

                    public partial class SalesDetailTable
                    {
                        public bool AddedByUser => true;
                    }
                    """
                ])
            .GeneratedType("SalesDetailTable")
            .GetProperty("AddedByUser")
            .Should().NotBeNull();
    }

    [Fact]
    public void 生成された行データ型は別ファイルのpartial定義と共にコンパイルできます()
    {
        GeneratedSourceCompiler
            .Compile(
                [
                    .. GeneratedCodeInspection.GenerateSources(
                        BasicStructureExcelFilePath),
                    """
                    namespace Generated;

                    public partial class SalesDetail
                    {
                        public bool AddedByUser => true;
                    }
                    """
                ])
            .GeneratedType("SalesDetail")
            .GetProperty("AddedByUser")
            .Should().NotBeNull();
    }
}
