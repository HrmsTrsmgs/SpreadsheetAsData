using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.CSharpIdentifier;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// コード生成設定から、生成コード上の名前を解決します。
/// </summary>
static class CodeGenerationOptionNames
{
    /// <summary>
    /// 単純な名前設定を優先して、生成コード上の名前を決定します。
    /// </summary>
    internal static string GeneratedName(
        this CodeGenerationOptions options,
        string sourceName) =>
        options.NameMappings.GetValueOrDefault(sourceName)
            ?? Identifier(sourceName);

    /// <summary>
    /// 文脈付き名前設定を優先して、生成コード上の名前を決定します。
    /// </summary>
    internal static string GeneratedName(
        this CodeGenerationOptions options,
        string contextName,
        string sourceName) =>
        options.NameMappings.GetValueOrDefault($"{contextName}.{sourceName}")
            ?? options.GeneratedName(sourceName);

    /// <summary>
    /// ブックスコープ定義名に対応する生成プロパティ名を決定します。
    /// </summary>
    internal static string BookDefinedName(
        this CodeGenerationOptions options,
        DefinedName definedName) =>
        options.GeneratedName("book", definedName.Name);

    /// <summary>
    /// ワークシートスコープ定義名に対応する生成プロパティ名を決定します。
    /// </summary>
    internal static string SheetDefinedName(
        this CodeGenerationOptions options,
        Worksheet sheet,
        DefinedName definedName) =>
        options.GeneratedName(sheet.Name, definedName.Name);

    /// <summary>
    /// Excelテーブル列に対応する生成プロパティ名を決定します。
    /// </summary>
    internal static string TableColumn(
        this CodeGenerationOptions options,
        Table table,
        TableColumn column) =>
        options.GeneratedName(table.Name, column.Name);
}
