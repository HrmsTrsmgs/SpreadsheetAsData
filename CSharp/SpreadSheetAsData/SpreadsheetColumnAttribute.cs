namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 型付きテーブルのプロパティへ対応付ける Excel テーブル列名を指定します。
/// </summary>
[AttributeUsage(
    AttributeTargets.Property,
    AllowMultiple = false,
    Inherited = true)]
public sealed class SpreadsheetColumnAttribute : Attribute
{
    /// <summary>
    /// 指定した Excel テーブル列名で属性を作成します。
    /// </summary>
    /// <param name="name">対応付ける Excel テーブル列名。</param>
    public SpreadsheetColumnAttribute(string name)
    {
        Name = name;
    }

    /// <summary>
    /// 対応付ける Excel テーブル列名を取得します。
    /// </summary>
    public string Name { get; }
}
