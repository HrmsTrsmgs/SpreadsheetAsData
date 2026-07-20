using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.WorkbookWrapperComponents;

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
            SourceFile(filePath, options, book)
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
}
