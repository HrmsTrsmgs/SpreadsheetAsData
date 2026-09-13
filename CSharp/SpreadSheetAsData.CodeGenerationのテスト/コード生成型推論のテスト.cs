using System.Reflection;
using FluentAssertions;
using Marimo.SpreadSheetAsData;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成型推論のテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\衝突なし\定義名.xlsx";
    const string IntegratedExcelFilePath = @"TestData\コード生成\統合.xlsx";

    [Theory]
    [InlineData("MainCell", typeof(string))]
    [InlineData("Total", typeof(double))]
    public void ブックスコープの単一セル定義名を現在値と同じ型のDataプロパティとして生成します(
        string propertyName,
        Type propertyType)
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedType("定義名Data")
            .GetProperty(propertyName);

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(propertyType);
    }

    [Fact]
    public void ブックスコープの複数セル定義名を二次元の値列挙となるDataプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedType("定義名Data")
            .GetProperty("MainRange");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            typeof(IEnumerable<IEnumerable<object?>>));
    }

    [Fact]
    public void シートローカルの単一セル定義名を現在値と同じ型のDataプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedType("定義名Data")
            .GetProperty("LocalCell");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(double));
    }

    [Fact]
    public void シートローカルの複数セル定義名を二次元の値列挙となるDataプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedType("定義名Data")
            .GetProperty("LocalRange");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(
            typeof(IEnumerable<IEnumerable<object?>>));
    }

    [Theory]
    [InlineData("MainCell")]
    [InlineData("MainRange")]
    [InlineData("LocalCell")]
    [InlineData("LocalRange")]
    public void 自動名前変換で対応する定義名にはSpreadSheetName属性を生成しません(
        string propertyName)
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath))
            .GeneratedType("定義名Data")
            .GetProperty(propertyName);

        tested.Should().NotBeNull();
        tested.GetCustomAttribute<SpreadSheetNameAttribute>()
            .Should().BeNull();
    }

    [Theory]
    [InlineData("SalesDetail")]
    [InlineData("ProductList")]
    public void Excelテーブルを生成行データの列挙となるDataプロパティとして生成します(
        string generatedName)
    {
        var generatedAssembly = GeneratedCodeInspection.AssemblyFrom(
            GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath));
        var rowType = generatedAssembly.GeneratedType(generatedName);
        var expectedType = typeof(IEnumerable<>).MakeGenericType(rowType);
        var tested = generatedAssembly
            .GeneratedType("BasicStructureData")
            .GetProperty(generatedName);

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(expectedType);
    }

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

    [Theory]
    [InlineData((double)int.MinValue, typeof(int))]
    [InlineData((double)int.MaxValue, typeof(int))]
    [InlineData((double)int.MinValue - 1, typeof(double))]
    [InlineData((double)int.MaxValue + 1, typeof(double))]
    public void 整数値の列はintに収まる場合だけintプロパティとし範囲外ならdoubleプロパティとして生成します(
        double value,
        Type propertyType)
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        var excelFilePath = temporaryFiles.Copy(BasicStructureExcelFilePath);
        using (var book = Workbook.Open(excelFilePath))
        {
            book.Tables["sales_detail"].Rows.First()["customer_id"].Value = value;
            book.Save();
        }

        var tested = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(propertyType);
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

    [Fact]
    public void 文字列値と空白を持つ列をstringプロパティとして生成します()
    {
        using var temporaryFiles = new TemporaryExcelFiles();
        var excelFilePath = temporaryFiles.Copy(@"TestData\テーブル.xlsx");
        using (var book = Workbook.Open(excelFilePath))
        {
            book.Tables["型付き行マッピング"].Rows.First()["文字列"].Value = null;
            book.Save();
        }

        var tested = GeneratedCodeInspection
            .AssemblyFrom(GeneratedCodeInspection.GenerateSources(excelFilePath))
            .GeneratedType("型付き行マッピング")
            .GetProperty("文字列");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(string));
    }

    [Fact]
    public void 真偽値だけを持つ列をboolプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedType("型付き行マッピング")
            .GetProperty("真偽値");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(bool));
    }

    [Fact]
    public void 整数値と空白を持つ列をnullableなintプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedType("空白数値マッピング")
            .GetProperty("数値");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(int?));
    }

    [Fact]
    public void 真偽値と空白を持つ列をnullableなboolプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedType("空白数値マッピング")
            .GetProperty("真偽値");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(bool?));
    }

    [Fact]
    public void 小数値と空白を持つ列をnullableなdoubleプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(@"TestData\テーブル.xlsx"))
            .GeneratedType("空白数値マッピング")
            .GetProperty("小数");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(double?));
    }

    [Fact]
    public void 列名と生成プロパティ名が一致する場合は列属性を生成しません()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(BasicStructureExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("Amount");

        tested.Should().NotBeNull();
        tested.GetCustomAttribute<SpreadSheetNameAttribute>()
            .Should().BeNull();
    }

    [Fact]
    public void 列名と生成プロパティ名が異なる場合は列属性に元のExcel列名を設定します()
    {
        var generatedProperty = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(IntegratedExcelFilePath))
            .GeneratedType("SalesDetail")
            .GetProperty("CustomerId");

        generatedProperty.Should().NotBeNull();

        var tested =
            generatedProperty.GetCustomAttribute<SpreadSheetNameAttribute>();

        tested.Should().NotBeNull();
        tested.Name.Should().Be("customer_id");
    }
}
