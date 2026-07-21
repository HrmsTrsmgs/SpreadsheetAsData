
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData;
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
            .. from column in
                table.TableDefinitionPart.Table.TableColumns
                    .Elements<Spreadsheet.TableColumn>()
                    .WithIndex()
               select new TableColumn(table, column.Value, column.Index)
        ];
    }

    /// <summary>
    /// 指定した位置の列定義を取得します。
    /// </summary>
    /// <param name="index">取得する列の 0 始まりの位置。</param>
    /// <returns>指定した位置の列定義。</returns>
    public TableColumn this[int index] =>
        (0..items.Length).Contains(index)
            ? items[index]
            : throw new ArgumentOutOfRangeException(nameof(index));

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
        ).SingleOrDefault()
            ?? throw new KeyNotFoundException();

    /// <summary>
    /// 指定した名前の列定義が存在するかどうかを返します。
    /// </summary>
    /// <param name="name">確認する列名。</param>
    /// <returns>指定した名前の列定義が存在する場合は true。</returns>
    public bool Contains(string name) =>
        (
            from item in items
            where item.Name == name
            select item
        ).Any();

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
