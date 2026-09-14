using FluentAssertions;
using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.WorkbookWrapperComponents;
using Xunit;

namespace Marimo.SpreadSheetAsData.CodeGeneration.Test;

public sealed class WorkbookWrapperComponentsのテスト : IDisposable
{
    const string BasicStructureExcelFilePath = @"TestData\コード生成\BasicStructure.xlsx";
    const string DefinedNamesExcelFilePath = @"TestData\コード生成\衝突なし\定義名.xlsx";

    readonly CodeGenerationOptions options = new();
    readonly Workbook basicBook;
    readonly Workbook definedNamesBook;

    [Theory]
    [InlineData("Workbook", "public partial class BasicStructureBook : global::Marimo.SpreadSheetAsData.Workbook")]
    [InlineData("Worksheet", "public partial class SalesDataSheet : global::Marimo.SpreadSheetAsData.Worksheet")]
    [InlineData("Table", "public TableTable(global::Marimo.SpreadSheetAsData.Table source)")]
    [InlineData("System", "Open(global::System.IO.Stream stream)")]
    public void SourceFileは同名の生成型に隠される参照を完全修飾します(string generatedName, string expected)
    {
        SourceFile(
            BasicStructureExcelFilePath,
            new() { NameMappings = new() { ["sales_detail"] = generatedName } },
            basicBook)
            .Should().Contain(expected);
    }

    [Theory]
    [InlineData("public partial class BasicStructureBook : Workbook")]
    [InlineData("public partial class SalesDataSheet : Worksheet")]
    [InlineData("public SalesDetailTable(Table source)")]
    [InlineData("Open(System.IO.Stream stream)")]
    public void SourceFileは同名の生成型がない参照を短い表記にします(string expected)
    {
        SourceFile(BasicStructureExcelFilePath, options, basicBook)
            .Should().Contain(expected);
    }

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
                using System.Collections.Generic;
                using System.Linq;
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

                    /// <summary>
                    /// 生成元のシート、テーブルとその列、両スコープの定義名を確認して、Excelブックを開きます。
                    /// </summary>
                    /// <exception cref="System.IO.InvalidDataException">生成元のシート、テーブルとその列、両スコープの定義名のいずれかが存在しないか、テーブルの所属シートが異なるか、単一セルの定義名が複数セルを参照しています。</exception>
                    public static new BasicStructureBook Open(string filePath) =>
                        ValidateStructure(new(filePath));

                    BasicStructureBook(System.IO.Stream stream) : base(stream)
                    {
                    }

                    /// <summary>
                    /// 生成元のシート、テーブルとその列、両スコープの定義名を確認して、Stream上のExcelブックを開きます。
                    /// </summary>
                    /// <exception cref="System.IO.InvalidDataException">生成元のシート、テーブルとその列、両スコープの定義名のいずれかが存在しないか、テーブルの所属シートが異なるか、単一セルの定義名が複数セルを参照しています。</exception>
                    public static new BasicStructureBook Open(System.IO.Stream stream) =>
                        ValidateStructure(new(stream));

                    /// <summary>
                    /// 生成元に対応するシート、テーブルとその列、両スコープの定義名を確認し、不一致時は開いたブックを破棄します。
                    /// </summary>
                    static BasicStructureBook ValidateStructure(BasicStructureBook book)
                    {
                        if (new string[] { "SalesData", "ProductMaster" }
                            .Any(sheetName => !book.Sheets.ContainsKey(sheetName))
                            || new (string SheetName, string TableName)[] { ("SalesData", "sales_detail"), ("ProductMaster", "ProductList") }
                                .Any(name => !book.Tables.Any(
                                    table => table.Name == name.TableName && table.Worksheet.Name == name.SheetName))
                            || new (string TableName, string ColumnName)[] { ("sales_detail", "customer_id"), ("sales_detail", "Amount"), ("sales_detail", "Description"), ("ProductList", "Id"), ("ProductList", "Name") }
                                .Any(column => !book.Tables[column.TableName].Columns.Contains(column.ColumnName))
                            || new (string Name, bool RequiresSingleCell)[] {  }
                                .Any(name => !book.DefinedNames.Any(
                                    definedName => definedName.Worksheet is null && definedName.Name == name.Name
                                        && (!name.RequiresSingleCell
                                            || definedName.Range.TopLeftCell == definedName.Range.BottomRightCell)))
                            || new (string SheetName, string Name, bool RequiresSingleCell)[] {  }
                                .Any(name => !book.DefinedNames.Any(
                                    definedName => definedName.Worksheet?.Name == name.SheetName && definedName.Name == name.Name
                                        && (!name.RequiresSingleCell
                                            || definedName.Range.TopLeftCell == definedName.Range.BottomRightCell))))
                        {
                            book.Dispose();
                            throw new System.IO.InvalidDataException();
                        }

                        return book;
                    }

                    /// <summary>
                    /// Excelブック全体のデータを読み込みます。
                    /// </summary>
                    public BasicStructureData Read() =>
                        base.Read<BasicStructureData>();

                    /// <summary>
                    /// Excelブック全体のデータを置換します。
                    /// </summary>
                    public void Replace(BasicStructureData data) =>
                        base.Replace(data);

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
                /// Excelブック「BasicStructure」のデータを表します。
                /// </summary>
                public partial class BasicStructureData
                {

                    /// <summary>
                    /// Excelテーブル「sales_detail」の行データを取得または設定します。
                    /// </summary>
                    public IEnumerable<SalesDetail> SalesDetail { get; set; }

                    /// <summary>
                    /// Excelテーブル「ProductList」の行データを取得または設定します。
                    /// </summary>
                    public IEnumerable<ProductList> ProductList { get; set; }
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
                    [SpreadSheetName("customer_id")]
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
        TableDeclaration(basicBook.Tables["sales_detail"], options)
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
                    [SpreadSheetName("customer_id")]
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
                    [SpreadSheetName("customer_id")]
                    public int CustomerId { get; set; }
                """);
    }

    [Fact]
    public void ColumnAttributeDeclarationは列名とプロパティ名が異なる場合にSpreadSheetName属性を生成します()
    {
        ColumnAttributeDeclaration(
                basicBook.Tables["sales_detail"].Columns["customer_id"],
                "CustomerId")
            .Should().Be(
                """
                    [SpreadSheetName("customer_id")]

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
                    /// 定義名「main_cell」が表すセルの値を取得または設定します。
                    /// </summary>
                    public dynamic MainCell
                    {
                        get => Cell["main_cell"].Value;
                        set => Cell["main_cell"].Value = value;
                    }
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
                    /// 定義名「main_range」が表すセル範囲の値を取得または設定します。
                    /// </summary>
                    public IEnumerable<IEnumerable<object?>> MainRange
                    {
                        get => Range["main_range"].Values;
                        set => Range["main_range"].Values = value;
                    }
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
                    /// ワークシート「sales_data」の定義名「local_cell」が表すセルの値を取得または設定します。
                    /// </summary>
                    public dynamic LocalCell
                    {
                        get => Cell["local_cell"].Value;
                        set => Cell["local_cell"].Value = value;
                    }
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
                    /// ワークシート「sales_data」の定義名「local_range」が表すセル範囲の値を取得または設定します。
                    /// </summary>
                    public IEnumerable<IEnumerable<object?>> LocalRange
                    {
                        get => Range["local_range"].Values;
                        set => Range["local_range"].Values = value;
                    }
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
