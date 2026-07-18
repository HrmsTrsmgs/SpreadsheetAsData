namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 整数範囲を読みやすく扱うための内部拡張です。
/// </summary>
static class RangeExtensions
{
    /// <summary>
    /// 先頭基準の終端排他範囲に、指定した値が含まれるかどうかを返します。
    /// </summary>
    /// <param name="range">確認する範囲。</param>
    /// <param name="value">範囲に含まれるか確認する値。</param>
    /// <returns>値が範囲に含まれる場合は true。</returns>
    /// <exception cref="NotSupportedException">末尾基準の範囲を指定した場合。</exception>
    internal static bool Contains(this Range range, int value)
    {
        if (range.Start.IsFromEnd || range.End.IsFromEnd)
        {
            throw new NotSupportedException();
        }

        return value >= range.Start.Value && value < range.End.Value;
    }

    /// <summary>
    /// 先頭基準の終端排他範囲に、指定した符号なし整数値が含まれるかどうかを返します。
    /// </summary>
    /// <param name="range">確認する範囲。</param>
    /// <param name="value">範囲に含まれるか確認する値。</param>
    /// <returns>値が範囲に含まれる場合は true。</returns>
    /// <exception cref="NotSupportedException">末尾基準の範囲を指定した場合。</exception>
    internal static bool Contains(this Range range, uint value) =>
        value <= int.MaxValue && range.Contains((int)value);
}
