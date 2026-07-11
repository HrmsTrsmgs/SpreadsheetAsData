
namespace Marimo.SpreadSheetAsData;
/// <summary>
/// ワークシート上の空白セルの値を表します。
/// </summary>
public struct BlankValue
{
    /// <summary>
    /// 空白値を空文字列に変換します。
    /// </summary>
    /// <param name="_">変換する空白値。</param>
    /// <returns>空文字列。</returns>
    public static implicit operator string(BlankValue _) => "";

    /// <summary>
    /// 空白値を数値の 0 に変換します。
    /// </summary>
    /// <param name="_">変換する空白値。</param>
    /// <returns>0。</returns>
    public static implicit operator double(BlankValue _) => .0;

    /// <summary>
    /// 空白値を表す文字列を返します。
    /// </summary>
    /// <returns>空白値を表す文字列。</returns>
    public override string ToString() => "{Blank}";
}
