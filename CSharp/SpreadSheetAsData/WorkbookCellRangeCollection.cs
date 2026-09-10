namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブックスコープの定義名、またはシート名付きの参照からセル範囲を取得します。
/// </summary>
public class WorkbookCellRangeCollection
{
    /// <summary>
    /// 定義名と参照先のシートを解決するブックです。
    /// </summary>
    readonly Workbook book;

    /// <summary>
    /// 同じ定義名から取得した範囲の同一性を保持します。
    /// </summary>
    readonly Dictionary<string, CellRange> namedRangeCache = [];

    /// <summary>
    /// ブックの範囲参照へのアクセスを構成します。
    /// </summary>
    /// <param name="book">範囲参照を解決するブック。</param>
    internal WorkbookCellRangeCollection(Workbook book) => this.book = book;

    /// <summary>
    /// 定義名、またはシート名付きA1形式の範囲参照からセル範囲を取得します。
    /// </summary>
    /// <param name="reference">解決する範囲参照。</param>
    /// <returns>指定した範囲参照が表すセル範囲。</returns>
    /// <exception cref="InvalidOperationException">参照先のワークシートを特定できない場合。</exception>
    public CellRange this[string reference]
    {
        get
        {
            if (namedRangeCache.TryGetValue(reference, out var cachedRange))
            {
                return cachedRange;
            }

            if (book.TryResolveNamedRange(reference, out var namedRange))
            {
                return namedRangeCache.GetValue(reference, () => namedRange);
            }

            if (!CellRangeReference.TryParse(reference, out var rangeReference))
            {
                throw new FormatException();
            }

            return new(
                book.Sheets[rangeReference.SheetName ?? throw new InvalidOperationException()],
                rangeReference.TopLeft,
                rangeReference.BottomRight);
        }
    }

    /// <summary>
    /// シート名とA1形式の左上セル、右下セルから範囲を取得します。
    /// </summary>
    /// <param name="sheetName">範囲が属するシート名。</param>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    /// <returns>指定したシート上のセル範囲。</returns>
    public CellRange this[string sheetName, string topLeft, string bottomRight] =>
        book.Sheets[sheetName].Range[topLeft, bottomRight];

    /// <summary>
    /// シート名と左上セル、右下セルのCellNameから範囲を取得します。
    /// </summary>
    /// <param name="sheetName">範囲が属するシート名。</param>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    /// <returns>指定したシート上のセル範囲。</returns>
    public CellRange this[string sheetName, CellName topLeft, CellName bottomRight] =>
        book.Sheets[sheetName].Range[topLeft, bottomRight];
}
