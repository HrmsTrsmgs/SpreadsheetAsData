using Marimo.SpreadSheetAsData;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// Excelブックから、SpreadsheetAsDataの型付きラッパーコードを生成します。
/// </summary>
public static class WorkbookWrapperGenerator
{
    /// <summary>
    /// 指定したExcelブックからC#ソースコードを生成します。
    /// </summary>
    /// <param name="filePath">生成元のExcelブックのパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>生成されたC#ソースコード。</returns>
    public static string[] GenerateSources(
        string filePath,
        Action<CodeGenerationOptions>? configure = null)
    {
        var options = new CodeGenerationOptions();
        configure?.Invoke(options);

        using var book = Workbook.Open(filePath);
        return
        [
            $$"""
            using Marimo.SpreadSheetAsData;

            namespace {{options.Namespace}};

            public partial class {{Path.GetFileNameWithoutExtension(filePath)}}Book : Workbook
            {
                public {{Path.GetFileNameWithoutExtension(filePath)}}Book() : base({{GenerateStringLiteral(filePath)}})
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
            """
        ];
    }

    /// <summary>
    /// 指定したExcelブックを解析し、コード生成前に検出できる問題を診断します。
    /// </summary>
    /// <param name="filePath">診断対象のExcelブックのパス。</param>
    /// <param name="configure">コード生成設定を変更する処理。</param>
    /// <returns>検出された診断情報。</returns>
    public static CodeGenerationDiagnostic[] GenerateDiagnostics(
        string filePath,
        Action<CodeGenerationOptions>? configure = null) =>
        [];

    static string SheetDeclaration(Worksheet sheet)
        => $$"""

        public partial class {{sheet.Name}}Sheet : Worksheet
        {
            public {{sheet.Name}}Sheet(Workbook book) : base(book, {{GenerateStringLiteral(sheet.Name)}})
            {
            }
        }
        """;

    static string TableDeclaration(Table table)
        => $$"""

        public partial class {{table.Name}}Table : Table<{{table.Name}}>
        {
            public {{table.Name}}Table(Table source) : base(source)
            {
            }
        }
        """;

    static string RowDeclaration(Table table)
        => $$"""

        public partial class {{table.Name}}
        {
        {{ForEach(
            from column in table.Columns
            select RowPropertyDeclaration(column))}}
        }
        """;

    static string RowPropertyDeclaration(TableColumn column) =>
        $"    public object? {column.Name} {{ get; set; }}";

    static string ForEach(IEnumerable<string> generatedBlocks) =>
        string.Join(Environment.NewLine, generatedBlocks);

    static string GenerateStringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
