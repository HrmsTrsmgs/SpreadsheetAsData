using System.Text.RegularExpressions;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// A1形式のセル範囲参照を表します。
/// </summary>
readonly partial struct CellRangeReference
{
    /// <summary>
    /// シート名を取得します。
    /// </summary>
    public string? SheetName { get; }

    /// <summary>
    /// 左上セル参照を取得します。
    /// </summary>
    public string TopLeft { get; }

    /// <summary>
    /// 右下セル参照を取得します。
    /// </summary>
    public string BottomRight { get; }

    /// <summary>
    /// 左上セル参照と右下セル参照からセル範囲参照を作成します。
    /// </summary>
    /// <param name="sheetName">シート名。</param>
    /// <param name="topLeft">左上セル参照。</param>
    /// <param name="bottomRight">右下セル参照。</param>
    CellRangeReference(string? sheetName, string topLeft, string bottomRight)
    {
        SheetName = sheetName;
        TopLeft = topLeft;
        BottomRight = bottomRight;
    }

    /// <summary>
    /// A1形式の範囲参照をセル範囲参照へ変換します。
    /// </summary>
    /// <param name="reference">A1形式の範囲参照。</param>
    /// <returns>変換したセル範囲参照。</returns>
    public static CellRangeReference Parse(string reference) =>
        TryParse(reference) ?? throw new NotImplementedException();

    [GeneratedRegex(
@"^(?:(?<sheet>[^!]*)!)?(?<startCell>[^!:]*):(?<endCell>[^!:]*)$")]
    private static partial Regex CellRangeReferencePattern();

    /// <summary>
    /// A1形式の範囲参照をセル範囲参照へ変換できる場合は変換します。
    /// </summary>
    /// <param name="reference">A1形式の範囲参照。</param>
    /// <returns>変換できた場合はセル範囲参照。変換できない場合はnull。</returns>
    public static CellRangeReference? TryParse(string reference)
    {
        var match = CellRangeReferencePattern().Match(reference);
        if (!match.Success)
        {
            return null;
        }

        var sheetName = match.Groups["sheet"].Success
            ? match.Groups["sheet"].Value
            : null;

        var startCellReference = match.Groups["startCell"].Value.Replace("$", "");
        var endCellReference = match.Groups["endCell"].Value.Replace("$", "");

        return new(
            sheetName,
            startCellReference,
            endCellReference);
    }
}
