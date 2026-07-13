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
        TryParse(reference, out var result)
            ? result
            : throw new NotImplementedException();

    [GeneratedRegex(
@"^(?:(?<sheet>[^!]*)!)?(?<startCell>[^!:]*):(?<endCell>[^!:]*)$")]
    private static partial Regex CellRangeReferencePattern();

    [GeneratedRegex(
@"^(?<sheet>[^!]*)!(?<cell>[^!:]*)$")]
    private static partial Regex SingleCellReferencePattern();

    [GeneratedRegex(
@"^\$?[A-Z]+\$?\d+$")]
    private static partial Regex CellReferencePattern();

    /// <summary>
    /// A1形式の範囲参照をセル範囲参照へ変換できる場合は変換します。
    /// </summary>
    /// <param name="reference">A1形式の範囲参照。</param>
    /// <param name="result">変換できた場合はセル範囲参照。変換できない場合は既定値。</param>
    /// <returns>変換できた場合はtrue。変換できない場合はfalse。</returns>
    public static bool TryParse(string reference, out CellRangeReference result)
    {
        result = default;

        var match = CellRangeReferencePattern().Match(reference);
        if (match.Success)
        {
            var sheetName = match.Groups["sheet"].Success
                ? match.Groups["sheet"].Value
                : null;

            var startCellReference = match.Groups["startCell"].Value;
            var endCellReference = match.Groups["endCell"].Value;
            if (!CellReferencePattern().IsMatch(startCellReference)
                || !CellReferencePattern().IsMatch(endCellReference))
            {
                return false;
            }

            var normalizedStartCellReference = startCellReference.Replace("$", "");
            var normalizedEndCellReference = endCellReference.Replace("$", "");

            try
            {
                CellName.Parse(normalizedStartCellReference);
                CellName.Parse(normalizedEndCellReference);
            }
            catch (FormatException)
            {
                return false;
            }

            result = new(
                sheetName,
                normalizedStartCellReference,
                normalizedEndCellReference);
            return true;
        }

        var singleCellMatch = SingleCellReferencePattern().Match(reference);
        if (!singleCellMatch.Success)
        {
            return false;
        }

        var singleCellReference = singleCellMatch.Groups["cell"].Value;
        if (!CellReferencePattern().IsMatch(singleCellReference))
        {
            return false;
        }

        result = new(
            singleCellMatch.Groups["sheet"].Value,
            singleCellReference.Replace("$", ""),
            singleCellReference.Replace("$", ""));
        return true;
    }
}
