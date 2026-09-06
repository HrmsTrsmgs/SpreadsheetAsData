namespace Marimo.SpreadSheetAsData;

/// <summary>
/// プロパティへ対応付ける Excel の定義名を指定します。
/// </summary>
[AttributeUsage(
    AttributeTargets.Property,
    AllowMultiple = false,
    Inherited = true)]
public sealed class SpreadsheetDefinedNameAttribute : Attribute
{
    /// <summary>
    /// 指定した Excel の定義名で属性を作成します。
    /// </summary>
    /// <param name="name">対応付ける Excel の定義名。</param>
    public SpreadsheetDefinedNameAttribute(string name)
    {
        Name = name;
    }

    /// <summary>
    /// 対応付ける Excel の定義名を取得します。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 定義名が属するワークシート名を取得または設定します。
    /// </summary>
    public string? WorksheetName { get; set; }
}
