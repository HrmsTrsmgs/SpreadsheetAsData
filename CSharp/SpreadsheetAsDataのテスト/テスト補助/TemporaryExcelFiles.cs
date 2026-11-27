using Marimo.SpreadsheetAsData.Testing;

namespace Marimo.SpreadsheetAsData.Test.テスト補助;

/// <summary>
/// 書き込みテストで、元のテストデータを変更しないための一時Excelファイルを管理します。
/// </summary>
sealed class TemporaryExcelFiles : IDisposable
{
    readonly TemporaryExcelFilePaths filePaths = new("SpreadsheetAsData.Tests");

    /// <summary>
    /// TestData配下のExcelファイルを一時ファイルへコピーします。
    /// </summary>
    /// <param name="sourceFileName">TestData配下のExcelファイル名。</param>
    /// <returns>コピー先の一時ファイルパス。</returns>
    internal string Copy(string sourceFileName)
    {
        var copiedPath = NewFilePath();

        using var source = File.Open(
            Path.Combine("TestData", sourceFileName),
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite);
        using var destination = File.Create(copiedPath);

        source.CopyTo(destination);

        return copiedPath;
    }

    /// <summary>
    /// 新しい一時Excelファイルパスを予約します。
    /// </summary>
    /// <returns>まだ存在しない一時Excelファイルパス。</returns>
    internal string NewFilePath() => filePaths.Reserve();

    public void Dispose() => filePaths.Dispose();
}
