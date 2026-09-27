namespace Marimo.SpreadsheetAsData;

/// <summary>
/// 型付きTableのセル値とプロパティ値を変換します。列の対応付けと変換失敗の診断は呼び出し側で扱います。
/// </summary>
static class TableValueConversion
{
    /// <summary>
    /// 元セル値を指定した型へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のセル値。</param>
    /// <param name="propertyType">変換先のプロパティ型。</param>
    /// <param name="converted">変換に成功した場合の値。</param>
    /// <returns>変換できた場合は true。</returns>
    internal static bool TryRead(
        object sourceValue,
        Type propertyType,
        out object? converted)
    {
        var valueType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;

        if (sourceValue is BlankValue
            && valueType != propertyType
            && (valueType == typeof(double)
                || valueType == typeof(int)
                || valueType == typeof(bool)))
        {
            converted = null;
            return true;
        }

        converted = (valueType, sourceValue) switch
        {
            ({ } type, _) when type == typeof(object) => sourceValue,
            ({ } type, double number) when type == typeof(int)
                && double.IsInteger(number)
                && number is >= int.MinValue and <= int.MaxValue => (int)number,
            ({ } type, double number) when type == typeof(double) => number,
            ({ } type, string text) when type == typeof(string) => text,
            ({ } type, BlankValue blank) when type == typeof(string) => (string)blank,
            ({ } type, BlankValue) when type == typeof(int) => 0,
            ({ } type, BlankValue blank) when type == typeof(double) => (double)blank,
            ({ } type, BlankValue) when type == typeof(bool) => false,
            ({ } type, bool boolean) when type == typeof(bool) => boolean,
            _ => null
        };

        return converted != null;
    }

    /// <summary>
    /// プロパティ値をセルへ直接設定できる値へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のプロパティ値。</param>
    /// <param name="converted">変換に成功した場合の値。</param>
    /// <returns>セルへ設定できる値に変換できた場合は true。</returns>
    internal static bool TryWrite(
        object? sourceValue,
        out object? converted)
    {
        if (sourceValue is null or "")
        {
            converted = null;
            return true;
        }

        converted = sourceValue switch
        {
            int number => (double)number,
            double number => number,
            string text => text,
            bool boolean => boolean,
            _ => null
        };

        return converted != null;
    }
}
