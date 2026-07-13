namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブック内の Excel テーブルを取得するコレクションを表します。
/// </summary>
public class TableCollection : IEnumerable<Table>
{
    /// <summary>
    /// ブック内の Excel テーブルを列挙します。
    /// </summary>
    /// <returns>Excel テーブルの列挙子。</returns>
    public IEnumerator<Table> GetEnumerator() =>
        throw new NotImplementedException();

    /// <summary>
    /// ブック内の Excel テーブルを列挙します。
    /// </summary>
    /// <returns>Excel テーブルの列挙子。</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        GetEnumerator();

    /// <summary>
    /// 指定した名前の Excel テーブルを取得します。
    /// </summary>
    /// <param name="name">取得する Excel テーブル名。</param>
    /// <returns>指定した名前の Excel テーブル。</returns>
    /// <exception cref="KeyNotFoundException">指定した名前の Excel テーブルが存在しない場合。</exception>
    public Table this[string name] =>
        throw new NotImplementedException();
}
