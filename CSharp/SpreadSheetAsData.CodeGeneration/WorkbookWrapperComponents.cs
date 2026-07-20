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
        $"    public {GeneratedName(sheet.Name, options)}Sheet {GeneratedName(sheet.Name, options)} => new(this);";

    /// <summary>
    /// Book型から指定Excelテーブル型を取得するプロパティ宣言を生成します。
    /// </summary>
    static string BookTablePropertyDeclaration(
        Table table,
        CodeGenerationOptions options) =>
        $"    public {GeneratedName(table.Name, options)}Table {GeneratedName(table.Name, options)} => new(Tables[{StringLiteral(table.Name)}]);";

    /// <summary>
    /// ワークシートを表す派生Sheet型の宣言を生成します。
    /// </summary>
    static string SheetDeclaration(
        Worksheet sheet,
        CodeGenerationOptions options)
        => $$"""

        public partial class {{GeneratedName(sheet.Name, options)}}Sheet : Worksheet
        {
            public {{GeneratedName(sheet.Name, options)}}Sheet(Workbook book) : base(book, {{StringLiteral(sheet.Name)}})
            {
            }
        {{ForEach(
            from definedName in SheetScopedDefinedNames(sheet)
            where IsSingleCellDefinedName(definedName)
            select SheetCellDefinedNamePropertyDeclaration(definedName, options))}}
        {{ForEach(
            from definedName in SheetScopedDefinedNames(sheet)
            where !IsSingleCellDefinedName(definedName)
            select SheetCellRangeDefinedNamePropertyDeclaration(definedName, options))}}
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
        CodeGenerationOptions options) =>
        $"    public object? {GeneratedName(table.Name, column.Name, options)} {{ get; set; }}";

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
        $"    public Cell {GeneratedName(definedName.Name, options)} => Cell[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// ブックスコープのセル範囲定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string BookCellRangeDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public CellRange {GeneratedName(definedName.Name, options)} => Range[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// ワークシートスコープの単一セル定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string SheetCellDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public Cell {GeneratedName(definedName.Name, options)} => Cell[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// ワークシートスコープのセル範囲定義名を取得するプロパティ宣言を生成します。
    /// </summary>
    static string SheetCellRangeDefinedNamePropertyDeclaration(
        DefinedName definedName,
        CodeGenerationOptions options) =>
        $"    public CellRange {GeneratedName(definedName.Name, options)} => Range[{StringLiteral(definedName.Name)}];";

    /// <summary>
    /// 文脈付き名前設定を優先して、生成コード上の名前を決定します。
    /// </summary>
    static string GeneratedName(
        string contextName,
        string sourceName,
        CodeGenerationOptions options) =>
        options.NameMappings.GetValueOrDefault($"{contextName}.{sourceName}")
            ?? GeneratedName(sourceName, options);

    /// <summary>
    /// 単純な名前設定を優先して、生成コード上の名前を決定します。
    /// </summary>
    static string GeneratedName(
        string sourceName,
        CodeGenerationOptions options) =>
        options.NameMappings.GetValueOrDefault(sourceName)
            ?? Identifier(sourceName);

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
