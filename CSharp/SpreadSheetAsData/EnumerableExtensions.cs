namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 列挙処理を読みやすくするための内部拡張です。
/// </summary>
static class EnumerableExtensions
{
    /// <summary>
    /// 各要素へ 0 始まりの位置を付けて列挙します。
    /// </summary>
    /// <typeparam name="T">列挙する要素の型。</typeparam>
    /// <param name="self">位置を付ける列挙。</param>
    /// <returns>元の要素と 0 始まりの位置を持つ列挙。</returns>
    internal static IEnumerable<(T Value, int Index)> WithIndex<T>(
        this IEnumerable<T> self) =>
        self.WithIndex(0);

    /// <summary>
    /// 各要素へ指定した開始位置からの位置を付けて列挙します。
    /// </summary>
    /// <typeparam name="T">列挙する要素の型。</typeparam>
    /// <param name="self">位置を付ける列挙。</param>
    /// <param name="startIndex">先頭要素へ付ける位置。</param>
    /// <returns>元の要素と指定した開始位置からの位置を持つ列挙。</returns>
    internal static IEnumerable<(T Value, int Index)> WithIndex<T>(
        this IEnumerable<T> self,
        int startIndex)
    {
        var index = startIndex;

        foreach (var value in self)
        {
            yield return (value, index);
            index = checked(index + 1);
        }
    }
}
