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
        var bookTypeName = GenerateTypeName(
            Path.GetFileNameWithoutExtension(filePath),
            "Book");
        var sheetDeclarations =
            from sheet in book.Sheets.Values
            select $$"""

            public partial class {{GenerateTypeName(sheet.Name, "Sheet")}} : Worksheet
            {
                public {{GenerateTypeName(sheet.Name, "Sheet")}}(Workbook book) : base(book, {{GenerateStringLiteral(sheet.Name)}})
                {
                }
            }
            """;
        var tableDeclarations =
            from table in book.Tables
            let rowTypeName = GenerateTypeName(table.Name, "")
            let tableTypeName = GenerateTypeName(table.Name, "Table")
            select $$"""

            public partial class {{tableTypeName}} : Table<{{rowTypeName}}>
            {
                public {{tableTypeName}}(Table source) : base(source)
                {
                }
            }
            """;
        var rowDeclarations =
            from table in book.Tables
            select $$"""

            public partial class {{GenerateTypeName(table.Name, "")}}
            {
            }
            """;

        return
        [
            $$"""
            using Marimo.SpreadSheetAsData;

            namespace {{options.Namespace}};

            public partial class {{bookTypeName}} : Workbook
            {
                public {{bookTypeName}}() : base({{GenerateStringLiteral(filePath)}})
                {
                }
            }
            {{string.Join(Environment.NewLine, sheetDeclarations)}}
            {{string.Join(Environment.NewLine, tableDeclarations)}}
            {{string.Join(Environment.NewLine, rowDeclarations)}}
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

    static string GenerateTypeName(string sourceName, string suffix) =>
        $"{sourceName}{suffix}";

    static string GenerateStringLiteral(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
}
