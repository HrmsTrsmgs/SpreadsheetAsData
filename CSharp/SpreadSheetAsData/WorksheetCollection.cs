using System.Collections;
using System.Diagnostics.CodeAnalysis;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ワークシートの読み取り専用コレクションです。
/// </summary>
public class WorksheetCollection : IReadOnlyList<Worksheet>, IReadOnlyDictionary<string, Worksheet>
{
    /// <summary>
    /// 受け取ったワークシート列挙を固定化した読み取り専用リストです。
    /// </summary>
    readonly Worksheet[] items;

    /// <summary>
    /// 指定したワークシート列挙からコレクションを作成します。
    /// </summary>
    /// <param name="collection">コレクションに含めるワークシート。</param>
    internal WorksheetCollection(IEnumerable<Worksheet> collection)
    {
        items = [.. collection];
    }

    /// <summary>
    /// 指定した名前のワークシートを取得します。
    /// </summary>
    /// <param name="sheetName">取得するワークシート名。</param>
    /// <returns>指定した名前のワークシート。</returns>
    /// <exception cref="KeyNotFoundException">指定した名前のワークシートが存在しません。</exception>
    public Worksheet this[string sheetName] =>
        (
            from sheet in items
            where sheet.Name == sheetName
            select sheet
        ).SingleOrDefault()
            ?? throw new KeyNotFoundException();

    /// <summary>
    /// 指定した位置のワークシートを取得します。
    /// </summary>
    /// <param name="index">取得するワークシートの 0 始まりの位置。</param>
    /// <returns>指定した位置のワークシート。</returns>
    public Worksheet this[int index] => items[index];

    /// <summary>
    /// コレクション内のワークシート数を取得します。
    /// </summary>
    public int Count => items.Length;

    /// <summary>
    /// ワークシート名の一覧を取得します。
    /// </summary>
    public IEnumerable<string> Keys =>
        from sheet in items
        select sheet.Name;

    /// <summary>
    /// ワークシートの一覧を取得します。
    /// </summary>
    public IEnumerable<Worksheet> Values => items;

    /// <summary>
    /// 指定した名前のワークシートが存在するかどうかを返します。
    /// </summary>
    /// <param name="key">確認するワークシート名。</param>
    /// <returns>存在する場合は true。</returns>
    public bool ContainsKey(string key) => Keys.Contains(key);

    /// <summary>
    /// 指定した名前のワークシートを取得します。
    /// </summary>
    /// <param name="key">取得するワークシート名。</param>
    /// <param name="value">取得したワークシート。存在しない場合は null。</param>
    /// <returns>指定した名前のワークシートが存在する場合は true。</returns>
    public bool TryGetValue(string key, [MaybeNullWhen(false)] out Worksheet value)
    {
        var sheet = (
            from item in items
            where item.Name == key
            select item
        ).SingleOrDefault();

        value = sheet;
        return sheet != null;
    }

    /// <summary>
    /// ワークシートを列挙する列挙子を返します。
    /// </summary>
    /// <returns>ワークシートを列挙する列挙子。</returns>
    public IEnumerator<Worksheet> GetEnumerator() =>
        ((IEnumerable<Worksheet>)items).GetEnumerator();

    /// <inheritdoc />
    IEnumerator<KeyValuePair<string, Worksheet>> IEnumerable<KeyValuePair<string, Worksheet>>.GetEnumerator() =>
        items.ToDictionary(_ => _.Name, _ => _).GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
