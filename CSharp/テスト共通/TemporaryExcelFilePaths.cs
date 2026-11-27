namespace Marimo.SpreadsheetAsData.Testing;

/// <summary>
/// テスト用の一時Excelファイルパスを予約し、終了時に作成済みのファイルを削除します。
/// </summary>
sealed class TemporaryExcelFilePaths(string directoryName) : IDisposable
{
    readonly List<string> filePaths = [];

    /// <summary>
    /// ファイルを作成せず、終了時に削除する対象としてパスを登録します。
    /// </summary>
    internal string Reserve()
    {
        var directory = Path.Combine(Path.GetTempPath(), directoryName);
        Directory.CreateDirectory(directory);

        var filePath = Path.Combine(directory, $"{Guid.NewGuid():N}.xlsx");
        filePaths.Add(filePath);
        return filePath;
    }

    public void Dispose()
    {
        foreach (var filePath in filePaths.Where(File.Exists))
        {
            File.Delete(filePath);
        }
    }
}
