using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ブック内のワークシートを表します。
/// </summary>
public class Worksheet
{
    /// <summary>
    /// ブックから作成されたワークシートだけが保持する親ブックです。
    /// </summary>
    readonly Workbook? book;

    /// <summary>
    /// ブックから作成されたワークシートだけが保持するシート名です。
    /// </summary>
    readonly string? name;

    /// <summary>
    /// ワークシートが持つコレクションを初期化します。
    /// </summary>
    Worksheet()
    {
        Cells = new CellCollection(this);
        Cell = new(this, resolvesWorksheetNames: true);
        Range = new(this);
    }

    /// <summary>
    /// 指定したブック内のワークシートを作成します。
    /// </summary>
    /// <param name="book">ワークシートが属するブック。</param>
    /// <param name="name">ワークシート名。</param>
    protected internal Worksheet(Workbook book, string name) : this()
    {
        this.book = book;
        this.name = name;
    }

    /// <summary>
    /// このワークシートが属するブックを取得します。
    /// </summary>
    public Workbook Book =>
        book ?? throw new InvalidOperationException();

    /// <summary>
    /// ワークシート名を取得します。
    /// </summary>
    public string Name =>
        name ?? throw new InvalidOperationException();

    /// <summary>
    /// ワークシート上のセルを取得するコレクションを取得します。
    /// </summary>
    public CellCollection Cells { get; }

    /// <summary>
    /// ワークシート上で有効なセル参照を解決するコレクションを取得します。
    /// </summary>
    public CellCollection Cell { get; }

    /// <summary>
    /// ワークシート上のセル範囲を取得するコレクションを取得します。
    /// </summary>
    public CellRangeCollection Range { get; }

    /// <summary>
    /// このワークシートに対応する Open XML のシート要素を取得します。
    /// </summary>
    internal Spreadsheet.Sheet SheetTag =>
        (
            from sheet in Book.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>()
            where sheet.Name == Name
            select sheet
        ).Single();

    /// <summary>
    /// このワークシートに対応する Open XML のワークシートパートを取得します。
    /// </summary>
    internal Packaging.WorksheetPart WorksheetPart =>
        Book.WorkbookPart.GetPartById(SheetTag.Id?.Value ?? throw new InvalidOperationException()) as Packaging.WorksheetPart
            ?? throw new InvalidOperationException();

    /// <summary>
    /// キャッシュに存在しないセルを、既存の Open XML セルまたは空白セルとして解決します。
    /// </summary>
    /// <param name="cellName">取得するセル参照。</param>
    /// <returns>指定したセル。</returns>
    internal Cell ResolveCell(CellName cellName)
    {
        var cellReference = cellName.ToString();
        var cellXml =
            from xml in WorksheetPart.Worksheet.Descendants<Spreadsheet.Cell>()
            where xml.CellReference == cellReference
            select xml;

        return cellXml.IsEmpty()
            ? new Cell(this, cellReference)
            : new Cell(this, cellXml.Single());
    }

    /// <summary>
    /// ワークシートスコープの定義名をセル範囲として解決します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <returns>定義名が表すセル範囲。</returns>
    internal CellRange ResolveNamedRange(string name)
    {
        var localSheetId = (uint)(
            from index in Enumerable.Range(0, Book.Sheets.Count)
            where Book.Sheets[index].Name == Name
            select index
        ).Single();

        return Book.ResolveNamedRange(name, localSheetId);
    }

    /// <summary>
    /// ワークシートスコープの定義名をセル範囲として解決できるか確認します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <param name="range">定義名が見つかった場合のセル範囲。</param>
    /// <returns>定義名を解決できた場合は true。</returns>
    internal bool TryResolveNamedRange(string name, out CellRange range)
    {
        var localSheetId = (uint)(
            from index in Enumerable.Range(0, Book.Sheets.Count)
            where Book.Sheets[index].Name == Name
            select index
        ).Single();

        return Book.TryResolveNamedRange(name, localSheetId, out range);
    }

    /// <summary>
    /// ワークシート名を返します。
    /// </summary>
    /// <returns>ワークシート名。</returns>
    public override string ToString() => Name;
}
