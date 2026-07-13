namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブック内の Excel テーブルを表します。
/// </summary>
public class Table
{
    /// <summary>
    /// Excel テーブル名を取得します。
    /// </summary>
    public string Name => throw new NotImplementedException();

    /// <summary>
    /// この Excel テーブルが属するワークシートを取得します。
    /// </summary>
    public Worksheet Worksheet => throw new NotImplementedException();

    /// <summary>
    /// ヘッダー行を含む Excel テーブル全体のセル範囲を取得します。
    /// </summary>
    public CellRange Range => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブルの列定義を取得します。
    /// </summary>
    public TableColumnCollection Columns => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブルのデータ行をワークシート上の順序で取得します。
    /// </summary>
    public IEnumerable<TableRow> Rows => throw new NotImplementedException();
}
