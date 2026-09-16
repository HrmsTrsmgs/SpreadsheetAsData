using FluentAssertions;
using Marimo.SpreadSheetAsData.CodeGeneration;
using Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class コード生成診断のテスト
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string ColumnNameCollisionExcelFilePath = @"TestData\コード生成\列名衝突.xlsx";
    const string BookMemberNameCollisionExcelFilePath = @"TestData\コード生成\Bookメンバー名衝突.xlsx";
    const string InvalidNameExcelFilePath = @"TestData\コード生成\無効名.xlsx";
    const string DefinedNamesWithoutCollisionsExcelFilePath = @"TestData\コード生成\衝突なし\定義名.xlsx";

    [Theory]
    [InlineData("SalesData", "Read")]
    [InlineData("sales_detail", "Read")]
    [InlineData("SalesData", "Open")]
    [InlineData("sales_detail", "Open")]
    [InlineData("SalesData", "Replace")]
    [InlineData("sales_detail", "Replace")]
    [InlineData("SalesData", "ValidateStructure")]
    [InlineData("sales_detail", "ValidateStructure")]
    [InlineData("SalesData", "Cell")]
    [InlineData("sales_detail", "Range")]
    public void シートやテーブルの生成プロパティ名がBookの既存メンバー名と衝突した場合に診断します(
        string sourceName,
        string generatedName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { [sourceName] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, [sourceName]));
    }

    [Theory]
    [InlineData("main_cell", "Read")]
    [InlineData("main_cell", "Open")]
    [InlineData("main_cell", "Replace")]
    [InlineData("main_cell", "ValidateStructure")]
    [InlineData("main_cell", "Cell")]
    [InlineData("main_range", "Range")]
    [InlineData("main_cell", "Tables")]
    [InlineData("main_cell", "DefinedNames")]
    [InlineData("main_cell", "Sheets")]
    [InlineData("main_cell", "Save")]
    [InlineData("main_cell", "SaveAs")]
    [InlineData("main_cell", "Close")]
    [InlineData("main_cell", "Dispose")]
    [InlineData("main_cell", "ReadTable")]
    public void ブックスコープの定義名がBookの既存メンバー名と衝突した場合に診断します(
        string sourceName,
        string generatedName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { [$"book.{sourceName}"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, [sourceName]));
    }

    [Theory]
    [InlineData("local_cell", "Cell")]
    [InlineData("local_range", "Range")]
    [InlineData("local_cell", "Book")]
    [InlineData("local_cell", "Name")]
    [InlineData("local_cell", "Cells")]
    [InlineData("local_cell", "ToString")]
    public void シートローカルの定義名がWorksheetの既存メンバー名と衝突した場合に診断します(
        string sourceName,
        string generatedName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { [$"sales_data.{sourceName}"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, [sourceName]));
    }

    [Theory(Skip = "定義名以外のBookプロパティにも、既存の予約名を共通して適用する段階で解除する。")]
    [InlineData("sales_detail", "Tables")]
    [InlineData("sales_detail", "DefinedNames")]
    [InlineData("sales_detail", "Sheets")]
    [InlineData("sales_detail", "Save")]
    [InlineData("sales_detail", "SaveAs")]
    [InlineData("sales_detail", "Close")]
    [InlineData("sales_detail", "Dispose")]
    [InlineData("sales_detail", "ReadTable")]
    [InlineData("SalesData", "Tables")]
    public void シートやテーブルのBookプロパティにも定義名と同じ既存メンバー名の衝突診断を適用します(
        string sourceName,
        string generatedName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { [sourceName] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, [sourceName]));
    }

    [Theory(Skip = "Sheetのテーブルプロパティにも、定義名と同じ既存メンバー名の予約を適用する段階で解除する。")]
    [InlineData("Name")]
    [InlineData("Cells")]
    [InlineData("ToString")]
    public void テーブルのSheetプロパティにも定義名と同じ既存メンバー名の衝突診断を適用します(string generatedName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["sales_detail"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, ["sales_detail"]));
    }

    [Theory]
    [InlineData("book.main_cell", "@Save", "Save", "main_cell")]
    [InlineData("sales_detail", "@Book", "Book", "sales_detail")]
    public void 定義名とテーブルの予約名診断はエスケープ表記が異なっても衝突を検出します(
        string mappingKey,
        string generatedName,
        string identifier,
        string sourceName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { [mappingKey] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, identifier, [sourceName]));
    }

    [Fact]
    public void 行データ型名にエスケープ表記があっても同じ識別子の列プロパティを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new()
                {
                    ["sales_detail"] = "@SalesDetail",
                    ["sales_detail.customer_id"] = "SalesDetail"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SalesDetail", ["customer_id"]));
    }

    [Fact]
    public void テーブルの生成プロパティにもWorksheetのBookプロパティ名との衝突診断を適用します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["sales_detail"] = "Book" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "Book", ["sales_detail"]));
    }

    [Theory]
    [InlineData("SalesData", "BasicStructureBook", "SalesData")]
    [InlineData("sales_detail", "BasicStructureBook", "sales_detail")]
    [InlineData("sales_detail", "SalesDataSheet", "sales_detail")]
    [InlineData("sales_detail.customer_id", "SalesDetail", "customer_id")]
    public void シートやテーブルや列の生成プロパティ名が所属する型名と衝突した場合に診断します(
        string mappingKey,
        string generatedName,
        string sourceName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { [mappingKey] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, [sourceName]));
    }

    [Theory]
    [InlineData("book.main_cell", "定義名Book", "main_cell")]
    [InlineData("sales_data.local_cell", "SalesDataSheet", "local_cell")]
    [InlineData("book.main_cell", "定義名Data", "main_cell")]
    [InlineData("sales_data.local_cell", "定義名Data", "local_cell")]
    [InlineData("sales_detail", "定義名Data", "sales_detail")]
    public void BookやSheetやDataの生成プロパティ名が所属する型名と衝突した場合に診断します(
        string mappingKey,
        string generatedName,
        string sourceName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { [mappingKey] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, [sourceName]));
    }

    [Fact]
    public void 自動変換後に同じ列プロパティ名となる場合にエラーを診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(ColumnNameCollisionExcelFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "CustomerId",
                    ["customer_id", "customer-id"]));
    }

    [Fact]
    public void 同じ生成型内の異なる種類のメンバー名が衝突した場合にも診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(BookMemberNameCollisionExcelFilePath)
            .Should().Contain(it => it.IsError
                && it.GeneratedName == "SalesData"
                && it.SourceNames.Contains("sales_data")
                && it.SourceNames.Contains("sales-data"));
    }

    [Fact]
    public void 同じBook型のシートとテーブルの生成プロパティ名が衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                BasicStructureExcelFilePath,
                options => options.NameMappings = new() { ["SalesData"] = "SalesDetail" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "SalesDetail",
                    ["SalesData", "sales_detail"]));
    }

    [Theory]
    [InlineData("SalesData", "sales_data")]
    [InlineData("SalesDetail", "sales_detail")]
    public void 同じBook型の定義名と他の生成プロパティ名が衝突した場合に診断します(
        string generatedName,
        string otherSourceName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["book.main_cell"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, ["main_cell", otherSourceName]));
    }

    [Fact]
    public void 同じSheet型の定義名とテーブルの生成プロパティ名が衝突した場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["sales_data.local_cell"] = "SalesDetail" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "SalesDetail",
                    ["local_cell", "sales_detail"]));
    }

    [Theory]
    [InlineData("book.main_cell", "book.main_range", "main_cell", "main_range")]
    [InlineData("sales_data.local_cell", "sales_data.local_range", "local_cell", "local_range")]
    public void 同じ生成型の単一セルと範囲の定義名が同じプロパティ名になる場合に診断します(
        string cellMappingKey,
        string rangeMappingKey,
        string cellName,
        string rangeName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    [cellMappingKey] = "SharedValue",
                    [rangeMappingKey] = "SharedValue"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SharedValue", [cellName, rangeName]));
    }

    [Fact]
    public void 同じBook型で三つ以上の生成プロパティ名が衝突した場合にすべての生成元を診断に含めます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    ["book.main_cell"] = "SalesDetail",
                    ["sales_data"] = "SalesDetail"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SalesDetail", ["main_cell", "sales_data", "sales_detail"]));
    }

    [Fact]
    public void ブックとシートの定義名がDataへ平坦化されて同じプロパティ名になる場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    ["book.main_cell"] = "SharedValue",
                    ["sales_data.local_cell"] = "SharedValue"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SharedValue", ["main_cell", "local_cell"]));
    }

    [Fact]
    public void シートローカルの定義名が別シートのテーブルとData内で同じプロパティ名になる場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["sales_data.local_cell"] = "ProductList" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "ProductList", ["local_cell", "product_list"]));
    }

    [Fact]
    public void 異なるシートから同じSheet型名を生成する場合に診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["product_master"] = "SalesData" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SalesDataSheet", ["sales_data", "product_master"]));
    }

    [Fact]
    public void 異なるテーブルから同じTable型名と行データ型名を生成する場合に両方を診断します()
    {
        var tested = GeneratedCodeInspection.GenerateDiagnostics(
            DefinedNamesWithoutCollisionsExcelFilePath,
            options => options.NameMappings = new() { ["product_list"] = "SalesDetail" });

        tested.Should().ContainEquivalentOf(
            new CodeGenerationDiagnostic(true, "SalesDetailTable", ["sales_detail", "product_list"]));
        tested.Should().ContainEquivalentOf(
            new CodeGenerationDiagnostic(true, "SalesDetail", ["sales_detail", "product_list"]));
    }

    [Theory]
    [InlineData("定義名Book", "定義名.xlsx")]
    [InlineData("定義名Data", "定義名.xlsx")]
    [InlineData("SalesDataSheet", "sales_data")]
    [InlineData("ProductListTable", "product_list")]
    public void 行データ型名が別の生成型名と衝突した場合に両方の生成元を診断します(
        string generatedName,
        string otherSourceName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["sales_detail"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, generatedName, ["sales_detail", otherSourceName]));
    }

    [Theory]
    [InlineData("Shared\u200CValue")]
    [InlineData("@SharedValue")]
    public void 書式文字やエスケープ表記だけが異なるプロパティ名も同じ識別子として診断します(string equivalentName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    ["book.main_cell"] = equivalentName,
                    ["book.main_range"] = "SharedValue"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SharedValue", ["main_cell", "main_range"]));
    }

    [Theory]
    [InlineData("sales_data.local_cell", "sales_data.local_range", "local_cell", "local_range")]
    [InlineData("sales_detail.customer_id", "sales_detail.amount", "customer_id", "amount")]
    [InlineData("book.main_cell", "sales_data.local_cell", "main_cell", "local_cell")]
    public void 書式文字を除くと同じになる名前の診断を各生成先のプロパティに適用します(
        string firstMappingKey,
        string secondMappingKey,
        string firstSourceName,
        string secondSourceName)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    [firstMappingKey] = "Shared\u200CValue",
                    [secondMappingKey] = "SharedValue"
                })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SharedValue", [firstSourceName, secondSourceName]));
    }

    [Theory]
    [InlineData("@Cell", "Cell")]
    [InlineData("Re\u200Cad", "Read")]
    public void 表記が異なっても既存メンバーと同じ識別子になる場合に診断します(
        string generatedName,
        string identifier)
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["book.main_cell"] = generatedName })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, identifier, ["main_cell"]));
    }

    [Fact]
    public void 書式文字を除くと同じになる生成型名も衝突として診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new() { ["product_list"] = "Sales\u200CDetail" })
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(true, "SalesDetail", ["sales_detail", "product_list"]));
    }

    [Fact]
    public void 大文字小文字だけが異なる生成プロパティ名は衝突にはなりません()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    ["book.main_cell"] = "SharedValue",
                    ["book.main_range"] = "sharedValue"
                })
            .Should().BeEmpty();
    }

    [Fact]
    public void 別の行データ型の列プロパティ名が同じでも衝突にはなりません()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                DefinedNamesWithoutCollisionsExcelFilePath,
                options => options.NameMappings = new()
                {
                    ["sales_detail.customer_id"] = "SharedValue",
                    ["product_list.id"] = "SharedValue"
                })
            .Should().BeEmpty();
    }

    [Fact]
    public void NameMappingsで生成名を変更すると名前衝突を解消できます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                ColumnNameCollisionExcelFilePath,
                options => options.NameMappings["customer-id"] = "CustomerIdDash")
            .Should().BeEmpty();
    }

    [Fact]
    public void 文脈付きNameMappingsでDataへ平坦化する同名定義名の衝突を解消できます()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(
                @"TestData\コード生成\定義名.xlsx",
                options => options.NameMappings = new() { ["sales_data.total"] = "SheetTotal" })
            .Should().BeEmpty();
    }

    [Fact]
    public void 名前衝突を自動的な連番追加では解消しません()
    {
        GeneratedCodeInspection
            .GenerateSources(ColumnNameCollisionExcelFilePath)
            .TypeDeclarations()
            .Should().NotContain(it => it.Identifier.ValueText.Contains("CustomerId2"));
    }

    [Fact]
    public void 区切り文字だけのシート名はCSharp識別子を生成できないため診断します()
    {
        GeneratedCodeInspection
            .GenerateDiagnostics(InvalidNameExcelFilePath)
            .Should().ContainEquivalentOf(
                new CodeGenerationDiagnostic(
                    true,
                    "",
                    ["---"],
                    "---"));
    }

}
