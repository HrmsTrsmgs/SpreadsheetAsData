using Xunit;

namespace Marimo.SpreadsheetAsData.Test.テスト補助;

/// <summary>
/// 作業ディレクトリの変更が他のテストの相対パス解決へ影響しないよう、並列実行を禁止します。
/// </summary>
[CollectionDefinition(nameof(CurrentDirectoryCollection), DisableParallelization = true)]
public sealed class CurrentDirectoryCollection
{
}
