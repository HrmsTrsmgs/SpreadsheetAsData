namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブックスコープの定義名、またはシート名と座標から単一セルを取得します。
/// </summary>
public class WorkbookCellCollection
{
    /// <summary>
    /// 定義名を解決し、参照先のワークシートからセルを取得するためのブックです。
    /// </summary>
    readonly Workbook book;

    /// <summary>
    /// ブックの名前付きセルへのアクセスを構成します。
    /// </summary>
    /// <param name="book">定義名を解決するブック。</param>
    internal WorkbookCellCollection(Workbook book) => this.book = book;

    /// <summary>
    /// ブックスコープの定義名が参照する単一セルを取得します。
    /// </summary>
    /// <param name="name">セルを参照する定義名。</param>
    /// <returns>定義名が参照するセル。</returns>
    /// <exception cref="KeyNotFoundException">指定した定義名が存在しない場合。</exception>
    /// <exception cref="InvalidOperationException">定義名が複数セルを参照する場合。</exception>
    public Cell this[string name] => book.ResolveNamedRange(name).SingleCell;

    /// <summary>
    /// シート名と列番号、行番号からセルを取得します。
    /// </summary>
    /// <param name="sheetName">セルが属するシート名。</param>
    /// <param name="columnIndex">1始まりの列番号。</param>
    /// <param name="rowIndex">1始まりの行番号。</param>
    /// <returns>指定したシート上のセル。</returns>
    public Cell this[string sheetName, uint columnIndex, uint rowIndex] =>
        book.Sheets[sheetName].Cells[columnIndex, rowIndex];

    /// <summary>
    /// シート名とセル参照からセルを取得します。
    /// </summary>
    /// <param name="sheetName">セルが属するシート名。</param>
    /// <param name="cellName">取得するセル参照。</param>
    /// <returns>指定したシート上のセル。</returns>
    public Cell this[string sheetName, CellName cellName] =>
        book.Sheets[sheetName].Cells[cellName];
}
