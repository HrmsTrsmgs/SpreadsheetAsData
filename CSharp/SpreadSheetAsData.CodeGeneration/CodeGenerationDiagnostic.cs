namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// コード生成前に検出した名前衝突や無効名などの診断情報を表します。
/// </summary>
/// <param name="IsError">コード生成を継続できないエラーである場合は <see langword="true"/>。</param>
/// <param name="GeneratedName">診断対象になった生成後のC#名。</param>
/// <param name="SourceNames">同じ生成名に対応したExcel上の元名。</param>
/// <param name="InvalidSourceName">有効なC#名へ変換できなかったExcel上の元名。</param>
public sealed record CodeGenerationDiagnostic(
    bool IsError,
    string GeneratedName,
    string[] SourceNames,
    string? InvalidSourceName = null);
