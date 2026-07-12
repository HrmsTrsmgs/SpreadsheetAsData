
namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ワークシート上のセル範囲を表します。
/// </summary>
public class CellRange
{
    /// <summary>
    /// <see cref="ToString"/> で A1 形式の範囲を復元するための右下セル参照です。
    /// </summary>
    readonly string bottomRight;

    /// <summary>
    /// 名前付き範囲として取得された場合の名前です。
    /// </summary>
    readonly string? name;

    /// <summary>
    /// 範囲のセル解決に使用するワークシートです。
    /// </summary>
    readonly Worksheet? sheet;

    /// <summary>
    /// <see cref="ToString"/> で A1 形式の範囲を復元するための左上セル参照です。
    /// </summary>
    readonly string topLeft;

    /// <summary>
    /// 指定した左上セルと右下セルでセル範囲を作成します。
    /// </summary>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    public CellRange(string topLeft, string bottomRight) : this(null, topLeft, bottomRight)
    { }

    /// <summary>
    /// 指定したワークシート上の左上セルと右下セルでセル範囲を作成します。
    /// </summary>
    /// <param name="sheet">範囲が属するワークシート。</param>
    /// <param name="topLeft">範囲の左上セル参照。</param>
    /// <param name="bottomRight">範囲の右下セル参照。</param>
    /// <param name="name">名前付き範囲として取得された場合の名前。</param>
    internal CellRange(Worksheet? sheet, string topLeft, string bottomRight, string? name = null)
    {
        this.sheet = sheet;
        this.topLeft = topLeft;
        this.bottomRight = bottomRight;
        this.name = name;
    }

    /// <summary>
    /// 名前付き範囲として取得された場合の名前を取得します。
    /// </summary>
    public string? Name => name;

    /// <summary>
    /// 範囲の左上セルを取得します。
    /// </summary>
    public Cell TopLeftCell =>
        (sheet ?? throw new NotImplementedException()).Cells[topLeft];

    /// <summary>
    /// 範囲が単一セルを表す場合に、そのセルを取得します。
    /// </summary>
    internal Cell SingleCell =>
        topLeft == bottomRight
            ? TopLeftCell
            : throw new InvalidOperationException();

    /// <summary>
    /// 範囲の右下セルを取得します。
    /// </summary>
    public Cell BottomRightCell =>
        (sheet ?? throw new NotImplementedException()).Cells[bottomRight];

    /// <summary>
    /// A1 形式のセル範囲を返します。
    /// </summary>
    /// <returns>A1 形式のセル範囲。</returns>
    public override string ToString() => $"{topLeft}:{bottomRight}";
}
