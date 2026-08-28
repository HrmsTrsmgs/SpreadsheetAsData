namespace Marimo.SpreadSheetAsData.CodeGeneration.Test.テスト補助;

/// <summary>
/// コード生成の書き込みテストで、元のテストデータを変更しないための一時Excelファイルを管理します。
/// </summary>
sealed class TemporaryExcelFiles : IDisposable
{
    readonly List<string> filePaths = [];

    /// <summary>
    /// 指定したExcelファイルを一時ファイルへコピーします。
    /// </summary>
    /// <param name="sourcePath">コピー元のExcelファイルパス。</param>
    /// <returns>コピー先の一時ファイルパス。</returns>
    internal string Copy(string sourcePath)
    {
        var copiedPath = NewFilePath();

        File.Copy(sourcePath, copiedPath);

        return copiedPath;
    }

    string NewFilePath()
    {
        var directory = Path.Combine(Path.GetTempPath(), "SpreadSheetAsData.CodeGeneration.Tests");
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
