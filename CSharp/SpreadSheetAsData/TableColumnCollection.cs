namespace Marimo.SpreadSheetAsData;

using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

/// <summary>
/// Excel テーブル内の列定義を取得するコレクションを表します。
/// </summary>
public class TableColumnCollection : IReadOnlyList<TableColumn>
{
    /// <summary>
    /// 定義順に固定した Excel テーブル列です。
    /// </summary>
    readonly TableColumn[] items;

    /// <summary>
    /// 指定した Excel テーブルの列コレクションを作成します。
    /// </summary>
    /// <param name="table">列定義を取得する Excel テーブル。</param>
    internal TableColumnCollection(Table table)
    {
        items = [
            .. from column in table.TableDefinitionPart.Table.TableColumns.Elements<Spreadsheet.TableColumn>()
               select new TableColumn(column)
        ];
    }

    /// <summary>
    /// 指定した位置の列定義を取得します。
    /// </summary>
    /// <param name="index">取得する列の 0 始まりの位置。</param>
    /// <returns>指定した位置の列定義。</returns>
    public TableColumn this[int index] =>
        items[index];

    /// <summary>
    /// 指定した名前の列定義を取得します。
    /// </summary>
    /// <param name="name">取得する列名。</param>
    /// <returns>指定した名前の列定義。</returns>
    /// <exception cref="KeyNotFoundException">指定した名前の列が存在しない場合。</exception>
    public TableColumn this[string name] =>
        (
            from item in items
            where item.Name == name
            select item
        ).Single();

    /// <summary>
    /// Excel テーブル内の列数を取得します。
    /// </summary>
    public int Count => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブル内の列定義を列挙します。
    /// </summary>
    /// <returns>列定義の列挙子。</returns>
    public IEnumerator<TableColumn> GetEnumerator() =>
        ((IEnumerable<TableColumn>)items).GetEnumerator();

    /// <summary>
    /// Excel テーブル内の列定義を列挙します。
    /// </summary>
    /// <returns>列定義の列挙子。</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
