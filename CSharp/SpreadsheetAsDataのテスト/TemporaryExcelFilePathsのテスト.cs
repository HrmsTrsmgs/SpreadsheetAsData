using FluentAssertions;
using Marimo.SpreadsheetAsData.Testing;
using Xunit;

namespace Marimo.SpreadsheetAsData.Test;

public sealed class TemporaryExcelFilePathsのテスト
{
    const string DirectoryName = "SpreadsheetAsData.TemporaryExcelFilePaths.Tests";

    [Fact]
    public void 予約は同じディレクトリにファイルを作成せず異なるExcelファイルパスを返します()
    {
        using var tested = new TemporaryExcelFilePaths(DirectoryName);
        var filePath = tested.Reserve();

        Path.GetDirectoryName(filePath).Should().Be(Path.Combine(Path.GetTempPath(), DirectoryName));
        Path.GetExtension(filePath).Should().Be(".xlsx");
        File.Exists(filePath).Should().BeFalse();
        tested.Reserve().Should().NotBe(filePath);
    }

    [Fact]
    public void 破棄時は自身の予約から作成されたファイルだけを削除します()
    {
        using var other = new TemporaryExcelFilePaths(DirectoryName);
        var otherFilePath = other.Reserve();
        File.WriteAllText(otherFilePath, string.Empty);

        string filePath;
        using (var tested = new TemporaryExcelFilePaths(DirectoryName))
        {
            tested.Reserve();
            filePath = tested.Reserve();
            File.WriteAllText(filePath, string.Empty);
        }

        File.Exists(filePath).Should().BeFalse();
        File.Exists(otherFilePath).Should().BeTrue();
    }
}
