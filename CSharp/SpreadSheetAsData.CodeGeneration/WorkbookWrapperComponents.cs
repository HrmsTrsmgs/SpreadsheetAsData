using Marimo.SpreadSheetAsData;

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
        string namespaceName,
        Workbook book) =>
        $$"""
        using Marimo.SpreadSheetAsData;

        namespace {{namespaceName}};

        public partial class {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book : Workbook
        {
            public {{Identifier(Path.GetFileNameWithoutExtension(filePath))}}Book() : base({{StringLiteral(filePath)}})
            {
            }
        }
        {{ForEach(
            from sheet in book.Sheets.Values
            select SheetDeclaration(sheet))}}
        {{ForEach(
            from table in book.Tables
            select TableDeclaration(table))}}
        {{ForEach(
            from table in book.Tables
            select RowDeclaration(table))}}
        """;

    static string SheetDeclaration(Worksheet sheet)
        => $$"""

        public partial class {{Identifier(sheet.Name)}}Sheet : Worksheet
        {
            public {{Identifier(sheet.Name)}}Sheet(Workbook book) : base(book, {{StringLiteral(sheet.Name)}})
            {
            }
        }
        """;

    static string TableDeclaration(Table table)
        => $$"""

        public partial class {{Identifier(table.Name)}}Table : Table<{{Identifier(table.Name)}}>
        {
            public {{Identifier(table.Name)}}Table(Table source) : base(source)
            {
            }
        }
        """;

    static string RowDeclaration(Table table)
        => $$"""

        public partial class {{Identifier(table.Name)}}
        {
        {{ForEach(
            from column in table.Columns
            select RowPropertyDeclaration(column))}}
        }
        """;

    static string RowPropertyDeclaration(TableColumn column) =>
        $"    public object? {Identifier(column.Name)} {{ get; set; }}";

    static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);

    internal static string Identifier(string sourceName) =>
        sourceName.Length == 0
            ? sourceName
            : $"{char.ToUpperInvariant(sourceName[0])}{sourceName[1..]}";

    static string StringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
