namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブル内のデータ行を表します。
/// </summary>
public class TableRow
{
    /// <summary>
    /// この行が属する Excel テーブルを取得します。
    /// </summary>
    public Table Table => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブルのデータ行内での 0 始まりの行位置を取得します。
    /// </summary>
    public int Ordinal => throw new NotImplementedException();

    /// <summary>
    /// ワークシート上の 1 始まりの行番号を取得します。
    /// </summary>
    public uint WorksheetRowIndex => throw new NotImplementedException();

    /// <summary>
    /// 指定した列名に対応するセルを取得します。
    /// </summary>
    /// <param name="columnName">取得するセルの列名。</param>
    /// <returns>指定した列に対応するセル。</returns>
    /// <exception cref="KeyNotFoundException">指定した名前の列が存在しない場合。</exception>
    public Cell this[string columnName] =>
        throw new NotImplementedException();

    /// <summary>
    /// 指定した列定義に対応するセルを取得します。
    /// </summary>
    /// <param name="column">取得するセルの列定義。</param>
    /// <returns>指定した列に対応するセル。</returns>
    /// <exception cref="ArgumentException">指定した列が別の Excel テーブルに属している場合。</exception>
    public Cell this[TableColumn column] =>
        throw new NotImplementedException();

    /// <summary>
    /// 指定した列位置に対応するセルを取得します。
    /// </summary>
    /// <param name="columnOrdinal">取得するセルの 0 始まりの列位置。</param>
    /// <returns>指定した列に対応するセル。</returns>
    /// <exception cref="ArgumentOutOfRangeException">指定した列位置が範囲外の場合。</exception>
    public Cell this[int columnOrdinal] =>
        throw new NotImplementedException();
}
