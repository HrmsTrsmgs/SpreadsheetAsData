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

        /// <summary>
        /// Excelブック「{{Path.GetFileNameWithoutExtension(filePath)}}」を型付きで表します。
        /// </summary>
        public partial class {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book : Workbook
        {
            public {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book() : this({{StringLiteral(filePath)}})
            {
            }

            public {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book(string filePath) : base(filePath)
            {
            }

            public static new {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book Open(string filePath) =>
                new(filePath);
        {{ForEach([
            .. from definedName in BookScopedDefinedNames(book)
               where IsSingleCellDefinedName(definedName)
               select BookCellDefinedNamePropertyDeclaration(definedName, options),
            .. from definedName in BookScopedDefinedNames(book)
               where !IsSingleCellDefinedName(definedName)
               select BookCellRangeDefinedNamePropertyDeclaration(definedName, options),
            .. from sheet in book.Sheets.Values
               select SheetPropertyDeclaration(sheet, options),
            .. from table in book.Tables
               select BookTablePropertyDeclaration(table, options)
        ])}}
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
    internal static string SheetPropertyDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// ワークシート「{{sheet.Name}}」を取得します。
            /// </summary>
            public {{options.GeneratedName(sheet.Name)}}Sheet {{options.GeneratedName(sheet.Name)}} => new(this);
        """;

    /// <summary>
    /// Book型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string BookTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// Excelテーブル「{{table.Name}}」を取得します。
            /// </summary>
            public {{options.GeneratedName(table.Name)}}Table {{options.GeneratedName(table.Name)}} => new(Tables[{{StringLiteral(table.Name)}}]);
        """;

    /// <summary>
    /// ワークシートを表す派生Sheet型の宣言を生成します。
    /// </summary>
    internal static string SheetDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options)
        => $$"""

        /// <summary>
        /// ワークシート「{{sheet.Name}}」を型付きで表します。
        /// </summary>
        public partial class {{options.GeneratedName(sheet.Name)}}Sheet : Worksheet
        {
            public {{options.GeneratedName(sheet.Name)}}Sheet(Workbook book) : base(book, {{StringLiteral(sheet.Name)}})
            {
            }
        {{ForEach([
            .. from definedName in SheetScopedDefinedNames(sheet)
               where IsSingleCellDefinedName(definedName)
               select SheetCellDefinedNamePropertyDeclaration(sheet, definedName, options),
            .. from definedName in SheetScopedDefinedNames(sheet)
               where !IsSingleCellDefinedName(definedName)
               select SheetCellRangeDefinedNamePropertyDeclaration(sheet, definedName, options),
            .. from table in sheet.Book.Tables
               where table.Worksheet.Name == sheet.Name
               select SheetTablePropertyDeclaration(table, options)
        ])}}
        }
        """;

    /// <summary>
    /// Excelテーブルを表す派生Table型の宣言を生成します。
    /// </summary>
    internal static string TableDeclaration(Table table)
        => $$"""

        /// <summary>
        /// Excelテーブル「{{table.Name}}」を型付きで表します。
        /// </summary>
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
    internal static string RowDeclaration(
        Table table,
        CodeGenerationOptions options)
        => $$"""

        /// <summary>
        /// Excelテーブル「{{table.Name}}」の1行を表します。
        /// </summary>
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
    internal static string RowPropertyDeclaration(
        Table table,
        TableColumn column,
        CodeGenerationOptions options)
    {
        var propertyTypeName = ColumnPropertyTypeName(table, column);
        var propertyName = options.TableColumn(table, column);

        return
            $$"""

                /// <summary>
                /// Excel列「{{column.Name}}」の値を取得または設定します。
                /// </summary>
            {{ColumnAttributeDeclaration(column, propertyName)}}    public {{propertyTypeName}} {{propertyName}} { get; set; }{{PropertyInitializer(propertyTypeName)}}
            """;
    }

    /// <summary>
    /// 生成プロパティ名とExcel列名が一致しない場合に、既存の型付きTableマッピングへ列名を伝える属性を生成します。
    /// </summary>
    internal static string ColumnAttributeDeclaration(TableColumn column, string propertyName) =>
        column.Name == propertyName
            ? ""
            : $"    [SpreadsheetColumn({StringLiteral(column.Name)})]{Environment.NewLine}";

    /// <summary>
    /// 既存の型付きTableマッピングで読み込めるプロパティ型名を、列の値から決定します。
    /// </summary>
    internal static string ColumnPropertyTypeName(Table table, TableColumn column)
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
    internal static string PropertyInitializer(string propertyTypeName) =>
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
    internal static string BookCellDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// 定義名「{{definedName.Name}}」が表すセルを取得します。
            /// </summary>
            public Cell {{options.BookDefinedName(definedName)}} => Cell[{{StringLiteral(definedName.Name)}}];
        """;

    /// <summary>
    /// ブックスコープのセル範囲定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string BookCellRangeDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// 定義名「{{definedName.Name}}」が表すセル範囲を取得します。
            /// </summary>
            public CellRange {{options.BookDefinedName(definedName)}} => Range[{{StringLiteral(definedName.Name)}}];
        """;

    /// <summary>
    /// ワークシートスコープの単一セル定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetCellDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// ワークシート「{{sheet.Name}}」の定義名「{{definedName.Name}}」が表すセルを取得します。
            /// </summary>
            public Cell {{options.SheetDefinedName(sheet, definedName)}} => Cell[{{StringLiteral(definedName.Name)}}];
        """;

    /// <summary>
    /// ワークシートスコープのセル範囲定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetCellRangeDefinedNamePropertyDeclaration(
        Worksheet sheet,
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// ワークシート「{{sheet.Name}}」の定義名「{{definedName.Name}}」が表すセル範囲を取得します。
            /// </summary>
            public CellRange {{options.SheetDefinedName(sheet, definedName)}} => Range[{{StringLiteral(definedName.Name)}}];
        """;

    /// <summary>
    /// Sheet型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    internal static string SheetTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $$"""

            /// <summary>
            /// Excelテーブル「{{table.Name}}」を取得します。
            /// </summary>
            public {{options.GeneratedName(table.Name)}}Table {{options.GeneratedName(table.Name)}} => new(Book.Tables[{{StringLiteral(table.Name)}}]);
        """;

    /// <summary>
    /// 複数のテンプレート部品を、生成ソース上の行単位で連結します。
    /// </summary>
    internal static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);

    /// <summary>
    /// 生成コード内へ埋め込む文字列リテラルを作ります。
    /// </summary>
    internal static string StringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
