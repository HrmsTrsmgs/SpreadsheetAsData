namespace Marimo.SpreadSheetAsData;

using Packaging = DocumentFormat.OpenXml.Packaging;

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
        new(this);

    /// <summary>
    /// Excel テーブルのデータ行をワークシート上の順序で取得します。
    /// </summary>
    public IEnumerable<TableRow> Rows =>
        from rowOffset in Enumerable.Range(0, DataRowCount)
        select new TableRow(FirstDataRowIndex + (uint)rowOffset);
}
