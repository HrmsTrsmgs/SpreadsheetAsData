
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ブック内の Excel テーブルを表します。
/// </summary>
public class Table
{
    /// <summary>
    /// Excel テーブル定義を保持する Open XML パートです。
    /// </summary>
    internal Packaging.TableDefinitionPart TableDefinitionPart { get; }

    /// <summary>
    /// この Excel テーブルが属するワークシートです。
    /// </summary>
    readonly Worksheet worksheet;

    /// <summary>
    /// Excel テーブル全体のセル範囲参照です。
    /// </summary>
    readonly CellRangeReference rangeReference;

    /// <summary>
    /// 同じ列定義を同じ <see cref="TableColumn"/> インスタンスとして返す列コレクションです。
    /// </summary>
    readonly TableColumnCollection columns;

    /// <summary>
    /// 同じデータ行を同じ <see cref="TableRow"/> インスタンスとして返すための行一覧です。
    /// </summary>
    readonly TableRow[] rows;

    /// <summary>
    /// ヘッダー行を除いたデータ行数です。
    /// </summary>
    int DataRowCount =>
        (int)(rangeReference.BottomRight.RowIndex - rangeReference.TopLeft.RowIndex);

    /// <summary>
    /// データ行が始まるワークシート上の 1 始まりの行番号です。
    /// </summary>
    uint FirstDataRowIndex =>
        rangeReference.TopLeft.RowIndex + 1;

    /// <summary>
    /// 指定した Open XML テーブル定義からテーブルを作成します。
    /// </summary>
    /// <param name="tableDefinitionPart">テーブル定義を保持する Open XML パート。</param>
    /// <param name="worksheet">Excel テーブルが属するワークシート。</param>
    internal Table(Packaging.TableDefinitionPart tableDefinitionPart, Worksheet worksheet)
    {
        TableDefinitionPart = tableDefinitionPart;
        this.worksheet = worksheet;
        rangeReference = CellRangeReference.Parse(tableDefinitionPart.Table.Reference.ToString());
        columns = new(this);
        rows = [
            .. from rowOffset in Enumerable.Range(0, DataRowCount)
               select new TableRow(this, rowOffset, FirstDataRowIndex + (uint)rowOffset)
        ];
    }

    /// <summary>
    /// Excel テーブル名を取得します。
    /// </summary>
    public string Name =>
        TableDefinitionPart.Table.Name.ToString();

    /// <summary>
    /// この Excel テーブルが属するワークシートを取得します。
    /// </summary>
    public Worksheet Worksheet =>
        worksheet;

    /// <summary>
    /// ヘッダー行を含む Excel テーブル全体のセル範囲を取得します。
    /// </summary>
    public CellRange Range =>
        worksheet.Range[rangeReference.TopLeft, rangeReference.BottomRight];

    /// <summary>
    /// Excel テーブルの列定義を取得します。
    /// </summary>
    public TableColumnCollection Columns =>
        columns;

    /// <summary>
    /// Excel テーブルのデータ行をワークシート上の順序で取得します。
    /// </summary>
    public IEnumerable<TableRow> Rows =>
        rows;

    /// <summary>
    /// Excel テーブルの各データ行を <typeparamref name="T"/> へ対応付けて列挙します。
    /// </summary>
    /// <typeparam name="T">各データ行を対応付ける型。</typeparam>
    /// <returns>型付き行の列挙。</returns>
    public IEnumerable<T> Enumerate<T>() =>
        new Table<T>(this);

    /// <summary>
    /// テーブル行と列定義の交点にあるワークシートセルを取得します。
    /// </summary>
    /// <param name="row">取得するセルが属するテーブル行。</param>
    /// <param name="column">取得するセルが属するテーブル列。</param>
    /// <returns>行と列の交点にあるセル。</returns>
    internal Cell ResolveCell(TableRow row, TableColumn column) =>
        column.Table == this
            ? worksheet.Cells[
                new CellName(
                    rangeReference.TopLeft.ColumnIndex + (uint)column.Ordinal,
                    row.WorksheetRowIndex)]
            : throw new ArgumentException(null, nameof(column));
}
