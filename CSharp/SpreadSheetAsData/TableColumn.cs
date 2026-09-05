
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// Excel テーブル内の列定義を表します。
/// </summary>
public class TableColumn
{
    /// <summary>
    /// この列が属する Excel テーブルです。
    /// </summary>
    readonly Table table;

    /// <summary>
    /// Open XML の列定義です。
    /// </summary>
    readonly Spreadsheet.TableColumn xml;

    /// <summary>
    /// Excel テーブル内での 0 始まりの列位置です。
    /// </summary>
    readonly int ordinal;

    /// <summary>
    /// 指定した Open XML 列定義からテーブル列を作成します。
    /// </summary>
    /// <param name="table">列が属する Excel テーブル。</param>
    /// <param name="xml">列名などを保持する Open XML 列定義。</param>
    /// <param name="ordinal">Excel テーブル内での 0 始まりの列位置。</param>
    internal TableColumn(Table table, Spreadsheet.TableColumn xml, int ordinal)
    {
        this.table = table;
        this.xml = xml;
        this.ordinal = ordinal;
    }

    /// <summary>
    /// Excel テーブル内の列名を取得します。
    /// </summary>
    public string Name =>
        xml.Name.ToString();

    /// <summary>
    /// Excel テーブル内での 0 始まりの列位置を取得します。
    /// </summary>
    public int Ordinal =>
        ordinal;

    /// <summary>
    /// この列が属する Excel テーブルを取得します。
    /// </summary>
    public Table Table =>
        table;

    /// <summary>
    /// Excel テーブル内の列名を返します。
    /// </summary>
    /// <returns>Excel テーブル内の列名。</returns>
    public override string ToString() => Name;
}
