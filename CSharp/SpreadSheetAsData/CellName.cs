using System.Text.RegularExpressions;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// A1 形式のセル参照を表します。
/// </summary>
public partial struct CellName
{
    /// <summary>
    /// Excel ワークシートで使用できる最大行番号です。
    /// </summary>
    public const uint MaxRowIndex = 1048576;

    /// <summary>
    /// Excel ワークシートで使用できる最大列番号です。
    /// </summary>
    public const uint MaxColumnIndex = 16384;

    /// <summary>
    /// Excel の列名を 26 進数相当で扱うための基数です。
    /// </summary>
    const uint alphabetCount = 26;

    /// <summary>
    /// 生成された正規表現を、セル参照の検証用プロパティとして扱います。
    /// </summary>
    static Regex CellNamePattern => GeneratedCellNameRegex();

    /// <summary>
    /// 1 始まりの列番号を取得します。
    /// </summary>
    public uint ColumnIndex { get; private set; }

    /// <summary>
    /// 1 始まりの行番号を取得します。
    /// </summary>
    public uint RowIndex { get; private set; }

    /// <summary>
    /// A1 形式の文字列を列番号と行番号へ分解してセル参照を作成します。
    /// </summary>
    /// <param name="name">A1 形式のセル参照。</param>
    /// <exception cref="FormatException">文字列がA1形式でない、または使用可能範囲を超えています。</exception>
    CellName(string name)
    {
        var match = CellNamePattern.Match(name);
        if (!match.Success)
        {
            throw new FormatException();
        }
        ColumnIndex = GetColumnIndex(match.Groups["column"].Value);
        RowIndex = uint.Parse(match.Groups["row"].Value);

        if (RowIndex < 1 || MaxRowIndex < RowIndex || MaxColumnIndex < ColumnIndex)
        {
            throw new FormatException();
        }
    }

    /// <summary>
    /// 列番号と行番号からセル参照を作成します。
    /// </summary>
    /// <param name="columnIndex">1 始まりの列番号。</param>
    /// <param name="rowIndex">1 始まりの行番号。</param>
    /// <exception cref="FormatException">列番号または行番号が使用可能範囲を超えています。</exception>
    public CellName(uint columnIndex, uint rowIndex)
    {
        ColumnIndex = columnIndex;
        RowIndex = rowIndex;
        if (MaxRowIndex < RowIndex || MaxColumnIndex < ColumnIndex)
        {
            throw new FormatException();
        }
    }

    /// <summary>
    /// A1 形式の文字列をセル参照に変換します。
    /// </summary>
    /// <param name="name">A1 形式のセル参照。</param>
    /// <returns>変換したセル参照。</returns>
    /// <exception cref="FormatException">文字列がA1形式でない、または使用可能範囲を超えています。</exception>
    public static CellName Parse(string name) => new(name);

    /// <summary>
    /// 列名を取得します。
    /// </summary>
    public readonly string ColumnName => GetColumnName(ColumnIndex);

    /// <summary>
    /// A1 形式のセル参照を返します。
    /// </summary>
    /// <returns>A1 形式のセル参照。</returns>
    public override readonly string ToString() =>
        $"{GetColumnName(ColumnIndex)}{RowIndex}";

    /// <summary>
    /// Excel の列名を 26 進数相当として 1 始まりの列番号に変換します。
    /// </summary>
    /// <param name="columnNameChars">列名を構成する文字列。</param>
    /// <returns>1 始まりの列番号。</returns>
    static uint GetColumnIndex(IEnumerable<char> columnNameChars) =>
        columnNameChars.Count() switch
        {
            1 => (uint)(columnNameChars.Single() - 'A') + 1,
            _ => GetColumnIndex(columnNameChars.Take(columnNameChars.Count() - 1)) * alphabetCount
                      + GetColumnIndex(columnNameChars.Skip(columnNameChars.Count() - 1))
        };

    /// <summary>
    /// 1 始まりの列番号を Excel の列名へ再帰的に変換します。
    /// </summary>
    /// <param name="columnIndex">1 始まりの列番号。</param>
    /// <returns>Excel の列名。</returns>
    static string GetColumnName(uint columnIndex) =>
        (columnIndex <= alphabetCount) switch
        {
            true => ((char)('A' + columnIndex - 1)).ToString(),
            false => $"{GetColumnName((columnIndex - 1) / alphabetCount)}{GetColumnName((columnIndex - 1) % alphabetCount + 1)}"
        };

    /// <summary>
    /// <see cref="GeneratedRegexAttribute"/> で A1 形式の検証用正規表現を生成します。
    /// </summary>
    /// <returns>A1 形式のセル参照を検証する正規表現。</returns>
    [GeneratedRegex(@"^(?<column>[A-Z]+)(?<row>\d+)$")]
    private static partial Regex GeneratedCellNameRegex();
}
