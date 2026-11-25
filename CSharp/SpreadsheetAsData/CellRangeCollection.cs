
namespace Marimo.SpreadsheetAsData;
/// <summary>
/// ワークシート上のセル範囲を取得するコレクションです。
/// </summary>
public class CellRangeCollection
{
    /// <summary>
    /// 同じ範囲指定に対して同じ <see cref="CellRange"/> インスタンスを返すためのキャッシュです。
    /// </summary>
    readonly Dictionary<(CellName TopLeft, CellName BottomRight), CellRange> cache = [];

    /// <summary>
    /// 同じ名前参照に対して同じ <see cref="CellRange"/> インスタンスを返すためのキャッシュです。
    /// </summary>
    readonly NamedRangeCache namedRangeCache;

    /// <summary>
    /// セル範囲が属するワークシートです。
    /// </summary>
    readonly Worksheet sheet;

    /// <summary>
    /// 指定したワークシート上のセル範囲コレクションを作成します。
    /// </summary>
    /// <param name="sheet">範囲参照を解決するワークシート。</param>
    internal CellRangeCollection(Worksheet sheet)
    {
        this.sheet = sheet;
        namedRangeCache = new(sheet.FindNamedRange);
    }

    /// <summary>
    /// 左上セルと右下セルを指定してセル範囲を取得します。
    /// </summary>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    /// <returns>指定したセル範囲。</returns>
    public CellRange this[string topLeft, string bottomRight] =>
        this[
            CellName.Parse(topLeft),
            CellName.Parse(bottomRight)];

    /// <summary>
    /// 左上セルと右下セルを指定してセル範囲を取得します。
    /// </summary>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    /// <returns>指定したセル範囲。</returns>
    public CellRange this[CellName topLeft, CellName bottomRight] =>
        cache.GetValue(
            (topLeft, bottomRight),
            () => new CellRange(
                sheet,
                topLeft,
                bottomRight));

    /// <summary>
    /// A1形式または名前による範囲参照からセル範囲を取得します。
    /// </summary>
    /// <param name="reference">解決する範囲参照。</param>
    /// <returns>指定した範囲参照が表すセル範囲。</returns>
    public CellRange this[string reference]
    {
        get
        {
            if (namedRangeCache.Find(reference) is { } namedRange)
            {
                return namedRange;
            }

            var rangeReference = CellRangeReference.Parse(reference);

            if (rangeReference.SheetName != null)
            {
                return sheet.Book.Range[
                    rangeReference.SheetName,
                    rangeReference.TopLeft,
                    rangeReference.BottomRight];
            }

            return this[rangeReference.TopLeft, rangeReference.BottomRight];
        }
    }
}
