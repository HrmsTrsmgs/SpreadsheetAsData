namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブック内の Excel テーブルを取得するコレクションを表します。
/// </summary>
public class TableCollection : IEnumerable<Table>
{
    /// <summary>
    /// テーブルを取得する対象ブックです。
    /// </summary>
    readonly Workbook book;

    /// <summary>
    /// 指定したブック内の Excel テーブルを取得するコレクションを作成します。
    /// </summary>
    /// <param name="book">テーブルを取得する対象ブック。</param>
    internal TableCollection(Workbook book)
    {
        this.book = book;
    }

    /// <summary>
    /// ブック内の Excel テーブルを列挙します。
    /// </summary>
    /// <returns>Excel テーブルの列挙子。</returns>
    public IEnumerator<Table> GetEnumerator() =>
        book.WorkbookPart.WorksheetParts
            .SelectMany(it => it.TableDefinitionParts)
            .Select(it => new Table(it))
            .GetEnumerator();

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
        this.SingleOrDefault(it => it.Name == name)
            ?? throw new KeyNotFoundException();
}
