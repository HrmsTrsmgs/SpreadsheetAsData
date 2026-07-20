using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成定義名のテスト
{
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

    [Fact]
    public void ブックスコープの単一セル定義名をBookのCellプロパティとして生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("定義名Book", "MainCell")
            .Type
            .ToString()
            .Should()
            .Be("Cell");
    }

    [Fact(
        Skip =
            "ブックスコープの複数セル定義名をBookのCellRange取得プロパティとして生成するときに解除する。")]
    public void ブックスコープの複数セル定義名をBookのCellRangeプロパティとして生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("定義名Book", "MainRange")
            .Type
            .ToString()
            .Should()
            .Be("CellRange");
    }

    [Fact(
        Skip =
            "シートローカルの単一セル定義名をSheetのCell取得プロパティとして生成するときに解除する。")]
    public void シートローカルの単一セル定義名をSheetのCellプロパティとして生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("SalesDataSheet", "LocalCell")
            .Type
            .ToString()
            .Should()
            .Be("Cell");
    }

    [Fact(
        Skip =
            "シートローカルの複数セル定義名をSheetのCellRange取得プロパティとして生成するときに解除する。")]
    public void シートローカルの複数セル定義名をSheetのCellRangeプロパティとして生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("SalesDataSheet", "LocalRange")
            .Type
            .ToString()
            .Should()
            .Be("CellRange");
    }

    [Fact(
        Skip =
            "同じ定義名をブックスコープとシートローカルで別々のプロパティとして生成するときに解除する。")]
    public void ブックスコープとシートローカルで同じ定義名を区別して生成します()
    {
        var sources = GeneratedCodeInspection.GenerateSources(
                    DefinedNamesExcelFilePath);

        sources.PropertyDeclaration("定義名Book", "Total").Should().NotBeNull();
        sources.PropertyDeclaration("SalesDataSheet", "Total").Should().NotBeNull();
    }
}


