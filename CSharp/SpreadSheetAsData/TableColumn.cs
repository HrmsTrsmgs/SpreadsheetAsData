namespace Marimo.SpreadSheetAsData;

using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

/// <summary>
/// Excel テーブル内の列定義を表します。
/// </summary>
public class TableColumn
{
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
    /// <param name="xml">列名などを保持する Open XML 列定義。</param>
    /// <param name="ordinal">Excel テーブル内での 0 始まりの列位置。</param>
    internal TableColumn(Spreadsheet.TableColumn xml, int ordinal)
    {
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
    public Table Table => throw new NotImplementedException();
}
