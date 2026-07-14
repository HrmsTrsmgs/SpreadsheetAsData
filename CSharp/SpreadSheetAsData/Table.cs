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
    /// 指定した Open XML テーブル定義からテーブルを作成します。
    /// </summary>
    /// <param name="tableDefinitionPart">テーブル定義を保持する Open XML パート。</param>
    internal Table(Packaging.TableDefinitionPart tableDefinitionPart)
    {
        TableDefinitionPart = tableDefinitionPart;
    }

    /// <summary>
    /// Excel テーブル名を取得します。
    /// </summary>
    public string Name =>
        TableDefinitionPart.Table.Name.ToString();

    /// <summary>
    /// この Excel テーブルが属するワークシートを取得します。
    /// </summary>
    public Worksheet Worksheet => throw new NotImplementedException();

    /// <summary>
    /// ヘッダー行を含む Excel テーブル全体のセル範囲を取得します。
    /// </summary>
    public CellRange Range => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブルの列定義を取得します。
    /// </summary>
    public TableColumnCollection Columns => throw new NotImplementedException();

    /// <summary>
    /// Excel テーブルのデータ行をワークシート上の順序で取得します。
    /// </summary>
    public IEnumerable<TableRow> Rows => throw new NotImplementedException();
}
