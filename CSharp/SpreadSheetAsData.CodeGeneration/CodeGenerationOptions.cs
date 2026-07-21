namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// ExcelブックからC#ラッパーコードを生成するときの設定を表します。
/// </summary>
public sealed class CodeGenerationOptions
{
    /// <summary>
    /// 生成するC#型を配置する名前空間を取得または設定します。
    /// </summary>
    public string Namespace { get; set; } = "Generated";

    /// <summary>
    /// Excel上の名前または文脈付き名前から生成後のC#名への対応表を取得または設定します。
    /// </summary>
    public Dictionary<string, string> NameMappings { get; set; } = [];
}
