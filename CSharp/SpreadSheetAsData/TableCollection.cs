
using Packaging = DocumentFormat.OpenXml.Packaging;

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
    /// Open XML のテーブル定義に対応する Table オブジェクトを保持します。
    /// </summary>
    readonly Dictionary<Packaging.TableDefinitionPart, Table> cache = [];

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
        (
            from sheet in book.Sheets.Values
            from tableDefinitionPart in sheet.WorksheetPart.TableDefinitionParts
            select GetTable(sheet, tableDefinitionPart)
        ).GetEnumerator();

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
        (
            from table in this
            where table.Name == name
            select table
        ).SingleOrDefault()
            ?? throw new KeyNotFoundException();

    Table GetTable(Worksheet sheet, Packaging.TableDefinitionPart tableDefinitionPart) =>
        cache.GetValue(
            tableDefinitionPart,
            () => new(tableDefinitionPart, sheet));
}
