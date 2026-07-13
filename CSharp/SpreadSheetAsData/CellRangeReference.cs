using System.Diagnostics.CodeAnalysis;
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
    public CellName TopLeft { get; }

    /// <summary>
    /// 右下セル参照を取得します。
    /// </summary>
    public CellName BottomRight { get; }

    /// <summary>
    /// 左上セル参照と右下セル参照からセル範囲参照を作成します。
    /// </summary>
    /// <param name="sheetName">シート名。</param>
    /// <param name="topLeft">左上セル参照。</param>
    /// <param name="bottomRight">右下セル参照。</param>
    CellRangeReference(string? sheetName, CellName topLeft, CellName bottomRight)
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

        if (!CellRangeReferenceText.TryCreate(reference, out var referenceText))
        {
            return false;
        }

        if (!TryParseCellReference(referenceText.StartCellReference, out var topLeft)
            || !TryParseCellReference(referenceText.EndCellReference, out var bottomRight))
        {
            return false;
        }

        result = new(
            referenceText.SheetName,
            topLeft,
            bottomRight);
        return true;
    }

    static bool TryParseCellReference(string reference, out CellName result)
    {
        result = default;

        if (!CellReferencePattern().IsMatch(reference))
        {
            return false;
        }

        try
        {
            result = CellName.Parse(reference.Replace("$", ""));
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    /// <summary>
    /// 正規表現の一致結果を、範囲参照として扱いやすい文字列へ成形します。
    /// </summary>
    sealed partial class CellRangeReferenceText
    {
        readonly Match match;

        CellRangeReferenceText(Match match)
        {
            this.match = match;
        }

        /// <summary>
        /// シート名を取得します。
        /// </summary>
        public string? SheetName => OptionalGroupValue("sheet");

        /// <summary>
        /// 始点セルの参照文字列を取得します。
        /// </summary>
        public string StartCellReference => GroupValue("startCell");

        /// <summary>
        /// 終点セルの参照文字列を取得します。
        /// </summary>
        public string EndCellReference => OptionalGroupValue("endCell") ?? StartCellReference;

        /// <summary>
        /// 参照文字列を正規表現で読み取り、範囲参照の文字列要素として取得します。
        /// </summary>
        /// <param name="reference">A1形式の範囲参照。</param>
        /// <param name="result">読み取れた場合は範囲参照の文字列要素。</param>
        /// <returns>正規表現で読み取れた場合はtrue。</returns>
        public static bool TryCreate(
            string reference,
            [NotNullWhen(true)] out CellRangeReferenceText? result)
        {
            var match = CellRangeReferencePattern().Match(reference);

            if (!match.Success)
            {
                result = null;
                return false;
            }

            result = new(match);
            return true;
        }

        string GroupValue(string groupName) => match.Groups[groupName].Value;

        string? OptionalGroupValue(string groupName)
        {
            var group = match.Groups[groupName];
            return group.Success
                ? group.Value
                : null;
        }

        [GeneratedRegex(
@"^(?:(?<sheet>[^!]*)!)?(?<startCell>[^!:]*)(?::(?<endCell>[^!:]*))?$")]
        private static partial Regex CellRangeReferencePattern();
    }
}
