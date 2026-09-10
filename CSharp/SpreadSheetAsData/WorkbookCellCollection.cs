namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブックスコープの定義名から単一セルを取得します。
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
}
