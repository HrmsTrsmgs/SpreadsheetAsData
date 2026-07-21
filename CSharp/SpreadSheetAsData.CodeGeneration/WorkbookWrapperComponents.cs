using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.CSharpIdentifier;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// Workbookラッパー生成で使用するテンプレート部品です。
/// </summary>
static class WorkbookWrapperComponents
{
    /// <summary>
    /// 生成するC#ソースファイル全体を表すテンプレート部品です。
    /// </summary>
    internal static string SourceFile(
        string filePath,
        CodeGenerationOptions options,
        Workbook book) =>
        $$"""
        using Marimo.SpreadSheetAsData;

        namespace {{options.Namespace}};

        public partial class {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book : Workbook
        {
            public {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book() : base({{StringLiteral(filePath)}})
            {
            }
        {{ForEach(
            from definedName in BookScopedDefinedNames(book)
            where IsSingleCellDefinedName(definedName)
            select BookCellDefinedNamePropertyDeclaration(definedName, options))}}
        {{ForEach(
            from definedName in BookScopedDefinedNames(book)
            where !IsSingleCellDefinedName(definedName)
            select BookCellRangeDefinedNamePropertyDeclaration(definedName, options))}}
        {{ForEach(
            from sheet in book.Sheets.Values
            select SheetPropertyDeclaration(sheet, options))}}
        {{ForEach(
            from table in book.Tables
            select BookTablePropertyDeclaration(table, options))}}
        }
        {{ForEach(
            from sheet in book.Sheets.Values
            select SheetDeclaration(sheet, options))}}
        {{ForEach(
            from table in book.Tables
            select TableDeclaration(table))}}
        {{ForEach(
            from table in book.Tables
            select RowDeclaration(table, options))}}
        """;

    /// <summary>
    /// Book型から指定ワークシート型を取得するプロパティ宣言を生成します。
    /// </summary>
    static string SheetPropertyDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options) =>
        $"    public {options.GeneratedName(sheet.Name)}Sheet {options.GeneratedName(sheet.Name)} => new(this);";

    /// <summary>
    /// Book型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    static string BookTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $"    public {options.GeneratedName(table.Name)}Table {options.GeneratedName(table.Name)} => new(Tables[{StringLiteral(table.Name)}]);";

    /// <summary>
    /// ワークシートを表す派生Sheet型の宣言を生成します。
    /// </summary>
    static string SheetDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options)
        => $$"""

        public partial class {{options.GeneratedName(sheet.Name)}}Sheet : Worksheet
        {
            public {{options.GeneratedName(sheet.Name)}}Sheet(Workbook book) : base(book, {{StringLiteral(sheet.Name)}})
            {
            }
        {{ForEach(
            from definedName in SheetScopedDefinedNames(sheet)
            where IsSingleCellDefinedName(definedName)
            select SheetCellDefinedNamePropertyDeclaration(sheet, definedName, options))}}
        {{ForEach(
            from definedName in SheetScopedDefinedNames(sheet)
            where !IsSingleCellDefinedName(definedName)
            select SheetCellRangeDefinedNamePropertyDeclaration(sheet, definedName, options))}}
        {{ForEach(
            from table in sheet.Book.Tables
            where table.Worksheet.Name == sheet.Name
            select SheetTablePropertyDeclaration(table, options))}}
        }
        """;

    /// <summary>
    /// Excelテーブルを表す派生Table型の宣言を生成します。
    /// </summary>
    static string TableDeclaration(Table table)
        => $$"""

        public partial class {{Identifier(table.Name)}}Table : Table<{{Identifier(table.Name)}}>
        {
            public {{Identifier(table.Name)}}Table(Table source) : base(source)
            {
            }
        }
        """;

    /// <summary>
    /// Excelテーブルの1行を表す行データ型の宣言を生成します。
    /// </summary>
    static string RowDeclaration(
        Table table,
        CodeGenerationOptions options)
        => $$"""

        public partial class {{Identifier(table.Name)}}
        {
        {{ForEach(
            from column in table.Columns
            select RowPropertyDeclaration(table, column, options))}}
        }
        """;

    /// <summary>
    /// Excelテーブル列に対応する行データプロパティ宣言を生成します。
    /// </summary>
    static string RowPropertyDeclaration(
        Table table,
        TableColumn column,
        CodeGenerationOptions options)
    {
        var propertyTypeName = ColumnPropertyTypeName(table, column);
        var propertyName = options.TableColumn(table, column);

        return
            $$"""
            {{ColumnAttributeDeclaration(column, propertyName)}}    public {{propertyTypeName}} {{propertyName}} { get; set; }{{PropertyInitializer(propertyTypeName)}}
            """;
    }

    /// <summary>
    /// 生成プロパティ名とExcel列名が一致しない場合に、既存の型付きTableマッピングへ列名を伝える属性を生成します。
    /// </summary>
    static string ColumnAttributeDeclaration(TableColumn column, string propertyName) =>
        column.Name == propertyName
            ? ""
            : $"    [SpreadsheetColumn({StringLiteral(column.Name)})]{Environment.NewLine}";

    /// <summary>
    /// 既存の型付きTableマッピングで読み込めるプロパティ型名を、列の値から決定します。
    /// </summary>
    static string ColumnPropertyTypeName(Table table, TableColumn column)
    {
        var values = (
            from row in table.Rows
            select row[column].Value
        ).ToArray();

        if (values.All(it => it is string))
        {
            return "string";
        }

        if (values.All(it => it is double number && double.IsInteger(number)))
        {
            return "int";
        }

        if (values.All(it => it is double))
        {
            return "double";
        }

        return "object?";
    }

    /// <summary>
    /// 生成プロパティがコンパイル時の初期化警告を出さないための初期値を返します。
    /// </summary>
    static string PropertyInitializer(string propertyTypeName) =>
        propertyTypeName == "string"
            ? " = \"\";"
            : "";

    /// <summary>
    /// ブック全体から参照できる定義名だけを選びます。
    /// </summary>
    static IEnumerable<DefinedName> BookScopedDefinedNames(Workbook book) =>
        from definedName in book.DefinedNames
        where definedName.Worksheet == null
        select definedName;

    /// <summary>
    /// 指定ワークシートだけで参照できる定義名を選びます。
    /// </summary>
    static IEnumerable<DefinedName> SheetScopedDefinedNames(Worksheet sheet) =>
        from definedName in sheet.Book.DefinedNames
        where definedName.Worksheet?.Name == sheet.Name
        select definedName;

    /// <summary>
    /// 定義名が単一セルを指すかどうかを判定します。
    /// </summary>
    static bool IsSingleCellDefinedName(DefinedName definedName) =>
        definedName.Range.TopLeftCell == definedName.Range.BottomRightCell;

    /// <summary>
    /// ブックスコープの単一セル定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string BookCellDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public Cell {options.BookDefinedName(definedName)} => Cell[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// ブックスコープのセル範囲定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string BookCellRangeDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public CellRange {options.BookDefinedName(definedName)} => Range[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// ワークシートスコープの単一セル定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string SheetCellDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public Cell {options.SheetDefinedName(sheet, definedName)} => Cell[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// ワークシートスコープのセル範囲定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string SheetCellRangeDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public CellRange {options.SheetDefinedName(sheet, definedName)} => Range[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// Sheet型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    static string SheetTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $"    public {options.GeneratedName(table.Name)}Table {options.GeneratedName(table.Name)} => new(Book.Tables[{StringLiteral(table.Name)}]);";

    /// <summary>
    /// 複数のテンプレート部品を、生成ソース上の行単位で連結します。
    /// </summary>
    static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);

    /// <summary>
    /// 生成コード内へ埋め込む文字列リテラルを作ります。
    /// </summary>
    static string StringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
