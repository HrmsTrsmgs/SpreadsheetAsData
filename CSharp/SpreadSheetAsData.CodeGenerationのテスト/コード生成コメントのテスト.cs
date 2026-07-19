using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成コメントのテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

    [Fact(
        Skip =
            "Book型のXML summaryへ元ブック名を含む固定文言を生成するときに解除する。")]
    public void Book型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("BasicStructureBook")
            .SummaryText()
            .Should()
            .Be("Excelブック「BasicStructure」を型付きで表します。");
    }

    [Fact(
        Skip =
            "Sheet型のXML summaryへ元ワークシート名を含む固定文言を生成するときに解除する。")]
    public void Sheet型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDataSheet")
            .SummaryText()
            .Should()
            .Be("ワークシート「SalesData」を型付きで表します。");
    }

    [Fact(
        Skip =
            "Table型のXML summaryへ元Excelテーブル名を含む固定文言を生成するときに解除する。")]
    public void Table型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDetailTable")
            .SummaryText()
            .Should()
            .Be("Excelテーブル「sales_detail」を型付きで表します。");
    }

    [Fact(
        Skip =
            "行データ型のXML summaryへ元Excelテーブル名を含む固定文言を生成するときに解除する。")]
    public void 行データ型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .TypeDeclaration("SalesDetail")
            .SummaryText()
            .Should()
            .Be("Excelテーブル「sales_detail」の1行を表します。");
    }

    [Fact(
        Skip =
            "SheetプロパティのXML summaryへ元ワークシート名を含む固定文言を生成するときに解除する。")]
    public void Sheetプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .PropertyDeclaration("BasicStructureBook", "SalesData")
            .SummaryText()
            .Should()
            .Be("ワークシート「SalesData」を取得します。");
    }

    [Fact(
        Skip =
            "TableプロパティのXML summaryへ元Excelテーブル名を含む固定文言を生成するときに解除する。")]
    public void Tableプロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .PropertyDeclaration("BasicStructureBook", "SalesDetail")
            .SummaryText()
            .Should()
            .Be("Excelテーブル「sales_detail」を取得します。");
    }

    [Fact(
        Skip =
            "列プロパティのXML summaryへ元Excel列名を含む固定文言を生成するときに解除する。")]
    public void 列プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureExcelFilePath)
            .PropertyDeclaration("SalesDetail", "CustomerId")
            .SummaryText()
            .Should()
            .Be("Excel列「customer_id」の値を取得または設定します。");
    }

    [Fact(
        Skip =
            "ブックスコープ単一セル定義名プロパティのXML summaryへ固定文言を生成するときに解除する。")]
    public void ブックスコープの単一セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("定義名Book", "MainCell")
            .SummaryText()
            .Should()
            .Be("定義名「main_cell」が表すセルを取得します。");
    }

    [Fact(
        Skip =
            "ブックスコープ複数セル定義名プロパティのXML summaryへ固定文言を生成するときに解除する。")]
    public void ブックスコープの複数セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("定義名Book", "MainRange")
            .SummaryText()
            .Should()
            .Be("定義名「main_range」が表すセル範囲を取得します。");
    }

    [Fact(
        Skip =
            "シートローカル単一セル定義名プロパティのXML summaryへ固定文言を生成するときに解除する。")]
    public void シートローカルの単一セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("SalesDataSheet", "LocalCell")
            .SummaryText()
            .Should()
            .Be("ワークシート「sales_data」の定義名「local_cell」が表すセルを取得します。");
    }

    [Fact(
        Skip =
            "シートローカル複数セル定義名プロパティのXML summaryへ固定文言を生成するときに解除する。")]
    public void シートローカルの複数セル定義名プロパティのコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(DefinedNamesExcelFilePath)
            .PropertyDeclaration("SalesDataSheet", "LocalRange")
            .SummaryText()
            .Should()
            .Be("ワークシート「sales_data」の定義名「local_range」が表すセル範囲を取得します。");
    }
}


