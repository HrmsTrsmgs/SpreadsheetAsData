namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 値を再利用する辞書を補助します。
/// </summary>
static class DictionaryExtensions
{
    extension<TKey, TValue>(IDictionary<TKey, TValue> self)
        where TKey : notnull
    {
        /// <summary>
        /// 指定したキーに対応する値を返します。
        /// </summary>
        /// <param name="key">値を取得するキー。</param>
        /// <param name="valueFactory">値がまだ保存されていない場合に使用する作成処理。</param>
        /// <returns>指定したキーに対応する値。</returns>
        internal TValue GetValue(TKey key, Func<TValue> valueFactory)
        {
            if (!self.TryGetValue(key, out var value))
            {
                value = valueFactory();
                self[key] = value;
            }

            return value;
        }
    }
}
