namespace Marimo.SpreadsheetAsData;

/// <summary>
/// 一つのスコープで解決できた定義名だけを記憶し、同じ名前には同じ範囲を返します。
/// </summary>
sealed class NamedRangeCache(Func<string, CellRange?> resolve)
{
    readonly Dictionary<string, CellRange> ranges = [];

    internal CellRange? Find(string name)
    {
        if (ranges.TryGetValue(name, out var cached))
        {
            return cached;
        }

        var range = resolve(name);
        if (range is not null)
        {
            ranges.Add(name, range);
        }
        return range;
    }
}
