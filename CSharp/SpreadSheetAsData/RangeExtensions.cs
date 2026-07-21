namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 整数範囲を読みやすく扱うための内部拡張です。
/// </summary>
static class RangeExtensions
{
    extension(Range self)
    {
        /// <summary>
        /// 先頭基準の終端排他範囲に、指定した値が含まれるかどうかを返します。
        /// </summary>
        /// <param name="value">範囲に含まれるか確認する値。</param>
        /// <returns>値が範囲に含まれる場合は true。</returns>
        /// <exception cref="NotSupportedException">末尾基準の範囲を指定した場合。</exception>
        internal bool Contains(int value)
        {
            if (self.Start.IsFromEnd || self.End.IsFromEnd)
            {
                throw new NotSupportedException();
            }

            return value >= self.Start.Value && value < self.End.Value;
        }

        /// <summary>
        /// 先頭基準の終端排他範囲に、指定した符号なし整数値が含まれるかどうかを返します。
        /// </summary>
        /// <param name="value">範囲に含まれるか確認する値。</param>
        /// <returns>値が範囲に含まれる場合は true。</returns>
        /// <exception cref="NotSupportedException">末尾基準の範囲を指定した場合。</exception>
        internal bool Contains(uint value) =>
            value <= int.MaxValue && self.Contains((int)value);
    }
}
