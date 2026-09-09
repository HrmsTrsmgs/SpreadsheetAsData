namespace Marimo.SpreadSheetAsData;

/// <summary>
/// 自動名前対応と異なる Excel 上の名前をプロパティへ対応付けます。
/// </summary>
[AttributeUsage(
    AttributeTargets.Property,
    AllowMultiple = false,
    Inherited = true)]
public sealed class SpreadSheetNameAttribute : Attribute
{
    /// <summary>
    /// 指定した Excel 上の名前で属性を作成します。
    /// </summary>
    /// <param name="name">対応付ける Excel 上の名前。</param>
    public SpreadSheetNameAttribute(string name)
    {
        ArgumentNullException.ThrowIfNull(name);

        if (name == "")
        {
            throw new ArgumentException(null, nameof(name));
        }

        Name = name;
    }

    /// <summary>
    /// 対応付ける Excel 上の名前を取得します。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// シートローカル定義名が属するワークシート名を取得または設定します。
    /// </summary>
    public string? WorksheetName { get; set; }
}
