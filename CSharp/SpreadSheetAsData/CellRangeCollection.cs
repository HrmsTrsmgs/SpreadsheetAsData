
namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ワークシート上のセル範囲を取得するコレクションです。
/// </summary>
public class CellRangeCollection
{
    /// <summary>
    /// ブックスコープの名前参照を解決するためのブックです。
    /// </summary>
    readonly Workbook? book;

    /// <summary>
    /// 同じ範囲指定に対して同じ <see cref="CellRange"/> インスタンスを返すためのキャッシュです。
    /// </summary>
    readonly Dictionary<(string TopLeft, string BottomRight), CellRange> cache = new();

    /// <summary>
    /// セル範囲が属するワークシートです。
    /// </summary>
    readonly Worksheet? sheet;

    /// <summary>
    /// 親を持たないセル範囲コレクションを作成します。
    /// </summary>
    public CellRangeCollection()
    { }

    /// <summary>
    /// 指定したブック上のセル範囲コレクションを作成します。
    /// </summary>
    /// <param name="book">範囲参照を解決するブック。</param>
    internal CellRangeCollection(Workbook book)
    {
        this.book = book;
    }

    /// <summary>
    /// 指定したワークシート上のセル範囲コレクションを作成します。
    /// </summary>
    /// <param name="sheet">範囲参照を解決するワークシート。</param>
    internal CellRangeCollection(Worksheet sheet)
    {
        this.sheet = sheet;
    }

    /// <summary>
    /// 左上セルと右下セルを指定してセル範囲を取得します。
    /// </summary>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    /// <returns>指定したセル範囲。</returns>
    public CellRange this[string topLeft, string bottomRight]
    {
        get
        {
            var key = (topLeft, bottomRight);
            if (!cache.TryGetValue(key, out var range))
            {
                range = new CellRange(sheet, topLeft, bottomRight);
                cache[key] = range;
            }

            return range;
        }
    }

    /// <summary>
    /// A1形式または名前による範囲参照からセル範囲を取得します。
    /// </summary>
    /// <param name="reference">解決する範囲参照。</param>
    /// <returns>指定した範囲参照が表すセル範囲。</returns>
    public CellRange this[string reference]
    {
        get
        {
            if (CellRangeReference.TryParse(reference) is not { } rangeReference)
            {
                return book?.GetWorkbookNamedRange(reference)
                    ?? throw new NotImplementedException();
            }

            if (rangeReference.SheetName != null)
            {
                return book != null
                    ? new(book.Sheets[rangeReference.SheetName], rangeReference.TopLeft, rangeReference.BottomRight)
                    : throw new NotImplementedException();
            }

            return this[rangeReference.TopLeft, rangeReference.BottomRight];
        }
    }
}
