using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ワークシート上のセルを取得するコレクションです。
/// </summary>
public class CellCollection
{
    /// <summary>
    /// ワークシートスコープの名前参照を解決するかどうかを表します。
    /// </summary>
    readonly bool resolvesWorksheetNames;

    /// <summary>
    /// セル参照とワークシートスコープの名前参照を解決するためのワークシートです。
    /// </summary>
    readonly Worksheet sheet;

    /// <summary>
    /// 指定したワークシートのセルコレクションを作成します。
    /// </summary>
    /// <param name="sheet">対象のワークシート。</param>
    public CellCollection(Worksheet sheet)
    {
        this.sheet = sheet;
    }

    /// <summary>
    /// 指定したワークシートのセルコレクションを作成します。
    /// </summary>
    /// <param name="sheet">対象のワークシート。</param>
    /// <param name="resolvesWorksheetNames">ワークシートスコープの名前参照を解決する場合は true。</param>
    internal CellCollection(Worksheet sheet, bool resolvesWorksheetNames) : this(sheet)
    {
        this.resolvesWorksheetNames = resolvesWorksheetNames;
    }

    /// <summary>
    /// A1 形式のセル参照でセルを取得します。
    /// </summary>
    /// <param name="cellReference">A1 形式のセル参照。</param>
    /// <returns>指定したセル。</returns>
    public Cell this[string cellReference]
    {
        get
        {
            if (resolvesWorksheetNames)
            {
                try
                {
                    return sheet.ResolveCell(CellName.Parse(cellReference));
                }
                catch (FormatException)
                {
                    return sheet.ResolveNamedRange(cellReference).SingleCell;
                }
            }

            return sheet.ResolveCell(CellName.Parse(cellReference));
        }
    }

    /// <summary>
    /// 列番号と行番号でセルを取得します。
    /// </summary>
    /// <param name="columnIndex">1 始まりの列番号。</param>
    /// <param name="rowIndex">1 始まりの行番号。</param>
    /// <returns>指定したセル。</returns>
    public Cell this[uint columnIndex, uint rowIndex] =>
        sheet.ResolveCell(new CellName(columnIndex, rowIndex));

    /// <summary>
    /// セル参照でセルを取得します。
    /// </summary>
    /// <param name="cellName">取得するセル参照。</param>
    /// <returns>指定したセル。</returns>
    public Cell this[CellName cellName] => sheet.ResolveCell(cellName);
}
