namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excelブックに定義された名前を表します。
/// 名前付きセルや名前付き範囲を対象とし、Excelテーブル名は含みません。
/// </summary>
public sealed class DefinedName
{
    /// <summary>
    /// 定義名を作成します。
    /// </summary>
    /// <param name="name">Excel上の定義名。</param>
    /// <param name="worksheet">定義名が属するワークシート。ブックスコープの場合は null。</param>
    /// <param name="range">定義名が表すセル範囲。</param>
    internal DefinedName(string name, Worksheet? worksheet, CellRange range)
    {
        Name = name;
        Worksheet = worksheet;
        Range = range;
    }

    /// <summary>
    /// Excel上の定義名を取得します。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 定義名が属するワークシートを取得します。
    /// ブックスコープの定義名では null を返します。
    /// </summary>
    public Worksheet? Worksheet { get; }

    /// <summary>
    /// 定義名が表すセル範囲を取得します。
    /// </summary>
    public CellRange Range { get; }

    /// <summary>
    /// Excel 上の定義名を返します。
    /// </summary>
    /// <returns>Excel 上の定義名。</returns>
    public override string ToString() => Name;
}
