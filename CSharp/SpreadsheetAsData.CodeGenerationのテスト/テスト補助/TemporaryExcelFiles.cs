using Marimo.SpreadsheetAsData.Testing;

namespace Marimo.SpreadsheetAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// コード生成の書き込みテストで、元のテストデータを変更しないための一時Excelファイルを管理します。
/// </summary>
sealed class TemporaryExcelFiles : IDisposable
{
    readonly TemporaryExcelFilePaths filePaths = new("SpreadsheetAsData.CodeGeneration.Tests");

    /// <summary>
    /// 指定したExcelファイルを一時ファイルへコピーします。
    /// </summary>
    /// <param name="sourcePath">コピー元のExcelファイルパス。</param>
    /// <returns>コピー先の一時ファイルパス。</returns>
    internal string Copy(string sourcePath)
    {
        var copiedPath = filePaths.Reserve();

        File.Copy(sourcePath, copiedPath);

        return copiedPath;
    }

    public void Dispose() => filePaths.Dispose();
}
