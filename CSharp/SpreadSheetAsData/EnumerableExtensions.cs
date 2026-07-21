namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 列挙処理を読みやすくするための内部拡張です。
/// </summary>
static class EnumerableExtensions
{
    extension<T>(IEnumerable<T> self)
    {
        /// <summary>
        /// 各要素へ 0 始まりの位置を付けて列挙します。
        /// </summary>
        /// <returns>元の要素と 0 始まりの位置を持つ列挙。</returns>
        internal IEnumerable<(T Value, int Index)> WithIndex() =>
            self.WithIndex(0);

        /// <summary>
        /// 各要素へ指定した開始位置からの位置を付けて列挙します。
        /// </summary>
        /// <param name="startIndex">先頭要素へ付ける位置。</param>
        /// <returns>元の要素と指定した開始位置からの位置を持つ列挙。</returns>
        internal IEnumerable<(T Value, int Index)> WithIndex(int startIndex)
        {
            var index = startIndex;

            foreach (var value in self)
            {
                yield return (value, index);
                index = checked(index + 1);
            }
        }

        /// <summary>
        /// 列挙の先頭要素を取得できるかどうかを返します。
        /// </summary>
        /// <param name="value">取得できた先頭要素。取得できなかった場合は使用しない。</param>
        /// <returns>先頭要素を取得できた場合は true。</returns>
        internal bool TryGetFirst(out T value)
        {
            using var enumerator = self.GetEnumerator();

            if (enumerator.MoveNext())
            {
                value = enumerator.Current;
                return true;
            }

            value = default!;
            return false;
        }
    }
}
