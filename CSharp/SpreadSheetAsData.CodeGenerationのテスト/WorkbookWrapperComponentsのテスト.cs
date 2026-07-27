using FluentAssertions;
using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.WorkbookWrapperComponents;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class WorkbookWrapperComponentsのテスト : IDisposable
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\定義名.xlsx";

    readonly CodeGenerationOptions options = new();
    readonly Workbook basicBook;
    readonly Workbook definedNamesBook;

    public WorkbookWrapperComponentsのテスト()
    {
        basicBook = Workbook.Open(BasicStructureExcelFilePath);
        definedNamesBook = Workbook.Open(DefinedNamesExcelFilePath);
    }

    public void Dispose()
    {
        basicBook.Dispose();
        definedNamesBook.Dispose();
    }

    [Fact]
    public void SheetPropertyDeclarationはBook型のワークシートプロパティ宣言を生成します()
    {
        SheetPropertyDeclaration(basicBook.Sheets["SalesData"], options)
            .Should().Be(
                """

                    /// <summary>
                    /// ワークシート「SalesData」を取得します。
                    /// </summary>
                    public SalesDataSheet SalesData => new(this);
                """);
    }

    [Fact]
    public void SourceFileは生成するCSharpソース全体を生成します()
    {
        SourceFile(BasicStructureExcelFilePath, options, basicBook)
            .Should().Be(
                """
                using Marimo.SpreadSheetAsData;

                namespace Generated;

                /// <summary>
                /// Excelブック「BasicStructure」を型付きで表します。
                /// </summary>
                public partial class BasicStructureBook : Workbook
                {
                    public BasicStructureBook() : this("TestData\\コード生成\\BasicStructure.xlsx")
                    {
                    }

                    public BasicStructureBook(string filePath) : base(filePath)
                    {
                    }

                    public static new BasicStructureBook Open(string filePath) =>
                        new(filePath);

                    /// <summary>
                    /// ワークシート「SalesData」を取得します。
                    /// </summary>
                    public SalesDataSheet SalesData => new(this);

                    /// <summary>
                    /// ワークシート「ProductMaster」を取得します。
                    /// </summary>
                    public ProductMasterSheet ProductMaster => new(this);

                    /// <summary>
                    /// Excelテーブル「sales_detail」を取得します。
                    /// </summary>
                    public SalesDetailTable SalesDetail => new(Tables["sales_detail"]);

                    /// <summary>
                    /// Excelテーブル「ProductList」を取得します。
                    /// </summary>
                    public ProductListTable ProductList => new(Tables["ProductList"]);
                }

                /// <summary>
                /// ワークシート「SalesData」を型付きで表します。
                /// </summary>
                public partial class SalesDataSheet : Worksheet
                {
                    public SalesDataSheet(Workbook book) : base(book, "SalesData")
                    {
                    }

                    /// <summary>
                    /// Excelテーブル「sales_detail」を取得します。
                    /// </summary>
                    public SalesDetailTable SalesDetail => new(Book.Tables["sales_detail"]);
                }

                /// <summary>
                /// ワークシート「ProductMaster」を型付きで表します。
                /// </summary>
                public partial class ProductMasterSheet : Worksheet
                {
                    public ProductMasterSheet(Workbook book) : base(book, "ProductMaster")
                    {
                    }

                    /// <summary>
                    /// Excelテーブル「ProductList」を取得します。
                    /// </summary>
                    public ProductListTable ProductList => new(Book.Tables["ProductList"]);
                }

                /// <summary>
                /// Excelテーブル「sales_detail」を型付きで表します。
                /// </summary>
                public partial class SalesDetailTable : Table<SalesDetail>
                {
                    public SalesDetailTable(Table source) : base(source)
                    {
                    }
                }

                /// <summary>
                /// Excelテーブル「ProductList」を型付きで表します。
                /// </summary>
                public partial class ProductListTable : Table<ProductList>
                {
                    public ProductListTable(Table source) : base(source)
                    {
                    }
                }

                /// <summary>
                /// Excelテーブル「sales_detail」の1行を表します。
                /// </summary>
                public partial class SalesDetail
                {

                    /// <summary>
                    /// Excel列「customer_id」の値を取得または設定します。
                    /// </summary>
                    [SpreadsheetColumn("customer_id")]
                    public int CustomerId { get; set; }

                    /// <summary>
                    /// Excel列「Amount」の値を取得または設定します。
                    /// </summary>
                    public double Amount { get; set; }

                    /// <summary>
                    /// Excel列「Description」の値を取得または設定します。
                    /// </summary>
                    public string Description { get; set; } = "";
                }

                /// <summary>
                /// Excelテーブル「ProductList」の1行を表します。
                /// </summary>
                public partial class ProductList
                {

                    /// <summary>
                    /// Excel列「Id」の値を取得または設定します。
                    /// </summary>
                    public int Id { get; set; }

                    /// <summary>
                    /// Excel列「Name」の値を取得または設定します。
                    /// </summary>
                    public string Name { get; set; } = "";
                }
                """);
    }

    [Fact]
    public void BookTablePropertyDeclarationはBook型のExcelテーブルプロパティ宣言を生成します()
    {
        BookTablePropertyDeclaration(basicBook.Tables["sales_detail"], options)
            .Should().Be(
                """

                    /// <summary>
                    /// Excelテーブル「sales_detail」を取得します。
                    /// </summary>
                    public SalesDetailTable SalesDetail => new(Tables["sales_detail"]);
                """);
    }

    [Fact]
    public void SheetDeclarationはSheet型宣言を生成します()
    {
        SheetDeclaration(basicBook.Sheets["SalesData"], options)
            .Should().Be(
                """

                /// <summary>
                /// ワークシート「SalesData」を型付きで表します。
                /// </summary>
                public partial class SalesDataSheet : Worksheet
                {
                    public SalesDataSheet(Workbook book) : base(book, "SalesData")
                    {
                    }

                    /// <summary>
                    /// Excelテーブル「sales_detail」を取得します。
                    /// </summary>
                    public SalesDetailTable SalesDetail => new(Book.Tables["sales_detail"]);
                }
                """);
    }

    [Fact]
    public void SheetTablePropertyDeclarationはXMLコメント付きでSheet型のExcelテーブルプロパティ宣言を生成します()
    {
        SheetTablePropertyDeclaration(basicBook.Tables["sales_detail"], options)
            .Should().Be(
                """

                    /// <summary>
                    /// Excelテーブル「sales_detail」を取得します。
                    /// </summary>
                    public SalesDetailTable SalesDetail => new(Book.Tables["sales_detail"]);
                """);
    }

    [Fact]
    public void TableDeclarationはTable型宣言を生成します()
    {
        TableDeclaration(basicBook.Tables["sales_detail"])
            .Should().Be(
                """

                /// <summary>
                /// Excelテーブル「sales_detail」を型付きで表します。
                /// </summary>
                public partial class SalesDetailTable : Table<SalesDetail>
                {
                    public SalesDetailTable(Table source) : base(source)
                    {
                    }
                }
                """);
    }

    [Fact]
    public void RowDeclarationは行データ型宣言を生成します()
    {
        RowDeclaration(basicBook.Tables["sales_detail"], options)
            .Should().Be(
                """

                /// <summary>
                /// Excelテーブル「sales_detail」の1行を表します。
                /// </summary>
                public partial class SalesDetail
                {

                    /// <summary>
                    /// Excel列「customer_id」の値を取得または設定します。
                    /// </summary>
                    [SpreadsheetColumn("customer_id")]
                    public int CustomerId { get; set; }

                    /// <summary>
                    /// Excel列「Amount」の値を取得または設定します。
                    /// </summary>
                    public double Amount { get; set; }

                    /// <summary>
                    /// Excel列「Description」の値を取得または設定します。
                    /// </summary>
                    public string Description { get; set; } = "";
                }
                """);
    }

    [Fact]
    public void RowPropertyDeclarationは列プロパティ宣言を生成します()
    {
        RowPropertyDeclaration(
                basicBook.Tables["sales_detail"],
                basicBook.Tables["sales_detail"].Columns["customer_id"],
                options)
            .Should().Be(
                """

                    /// <summary>
                    /// Excel列「customer_id」の値を取得または設定します。
                    /// </summary>
                    [SpreadsheetColumn("customer_id")]
                    public int CustomerId { get; set; }
                """);
    }

    [Fact]
    public void ColumnAttributeDeclarationは列名とプロパティ名が異なる場合にSpreadsheetColumn属性を生成します()
    {
        ColumnAttributeDeclaration(
                basicBook.Tables["sales_detail"].Columns["customer_id"],
                "CustomerId")
            .Should().Be(
                """
                    [SpreadsheetColumn("customer_id")]

                """);
    }

    [Fact]
    public void ColumnAttributeDeclarationは列名とプロパティ名が一致する場合に空文字列を返します()
    {
        ColumnAttributeDeclaration(
                basicBook.Tables["sales_detail"].Columns["Amount"],
                "Amount")
            .Should().Be("");
    }

    [Theory]
    [InlineData("customer_id", "int")]
    [InlineData("Amount", "double")]
    [InlineData("Description", "string")]
    public void ColumnPropertyTypeNameは列値からプロパティ型名を生成します(
        string columnName,
        string typeName)
    {
        ColumnPropertyTypeName(
                basicBook.Tables["sales_detail"],
                basicBook.Tables["sales_detail"].Columns[columnName])
            .Should().Be(typeName);
    }

    [Theory]
    [InlineData("string", " = \"\";")]
    [InlineData("int", "")]
    public void PropertyInitializerはプロパティ型名から初期化子を生成します(
        string propertyTypeName,
        string initializer)
    {
        PropertyInitializer(propertyTypeName)
            .Should().Be(initializer);
    }

    [Fact]
    public void BookCellDefinedNamePropertyDeclarationはブックスコープ単一セル定義名プロパティ宣言を生成します()
    {
        BookCellDefinedNamePropertyDeclaration(
                definedNamesBook.DefinedNames.Single(it => it.Name == "main_cell"),
                options)
            .Should().Be(
                """

                    /// <summary>
                    /// 定義名「main_cell」が表すセルを取得します。
                    /// </summary>
                    public Cell MainCell => Cell["main_cell"];
                """);
    }

    [Fact]
    public void BookCellRangeDefinedNamePropertyDeclarationはブックスコープ複数セル定義名プロパティ宣言を生成します()
    {
        BookCellRangeDefinedNamePropertyDeclaration(
                definedNamesBook.DefinedNames.Single(it => it.Name == "main_range"),
                options)
            .Should().Be(
                """

                    /// <summary>
                    /// 定義名「main_range」が表すセル範囲を取得します。
                    /// </summary>
                    public CellRange MainRange => Range["main_range"];
                """);
    }

    [Fact]
    public void SheetCellDefinedNamePropertyDeclarationはシートローカル単一セル定義名プロパティ宣言を生成します()
    {
        SheetCellDefinedNamePropertyDeclaration(
                definedNamesBook.Sheets["sales_data"],
                definedNamesBook.DefinedNames.Single(it => it.Name == "local_cell"),
                options)
            .Should().Be(
                """

                    /// <summary>
                    /// ワークシート「sales_data」の定義名「local_cell」が表すセルを取得します。
                    /// </summary>
                    public Cell LocalCell => Cell["local_cell"];
                """);
    }

    [Fact]
    public void SheetCellRangeDefinedNamePropertyDeclarationはシートローカル複数セル定義名プロパティ宣言を生成します()
    {
        SheetCellRangeDefinedNamePropertyDeclaration(
                definedNamesBook.Sheets["sales_data"],
                definedNamesBook.DefinedNames.Single(it => it.Name == "local_range"),
                options)
            .Should().Be(
                """

                    /// <summary>
                    /// ワークシート「sales_data」の定義名「local_range」が表すセル範囲を取得します。
                    /// </summary>
                    public CellRange LocalRange => Range["local_range"];
                """);
    }

    [Fact]
    public void ForEachは生成ブロックを改行で連結します()
    {
        ForEach(["first", "second", "third"])
            .Should().Be(
                """
                first
                second
                third
                """);
    }

    [Fact]
    public void StringLiteralは生成コード内の文字列リテラルを生成します()
    {
        StringLiteral(@"C:\temp\""book.xlsx")
            .Should().Be(
                """""
                "C:\\temp\\\"book.xlsx"
                """"");
    }
}
