namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブル内の列定義を表します。
/// </summary>
public class TableColumn
{
    /// <summary>
    /// Excel テーブル内の列名を取得します。
    /// </summary>
    public string Name => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブル内での 0 始まりの列位置を取得します。
    /// </summary>
    public int Ordinal => throw new NotImplementedException();

    /// <summary>
    /// この列が属する Excel テーブルを取得します。
    /// </summary>
    public Table Table => throw new NotImplementedException();
}
