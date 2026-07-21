using Marimo.SpreadSheetAsData;
using static Marimo.SpreadSheetAsData.CodeGeneration.CSharpIdentifier;

namespace Marimo.SpreadSheetAsData.CodeGeneration;

/// <summary>
/// コード生成設定から、生成コード上の名前を解決します。
/// </summary>
static class CodeGenerationOptionNames
{
    extension(CodeGenerationOptions self)
    {
        /// <summary>
        /// 単純な名前設定を優先して、生成コード上の名前を決定します。
        /// </summary>
        internal string GeneratedName(string sourceName) =>
            self.NameMappings.GetValueOrDefault(sourceName)
                ?? Identifier(sourceName);

        /// <summary>
        /// 文脈付き名前設定を優先して、生成コード上の名前を決定します。
        /// </summary>
        internal string GeneratedName(string contextName, string sourceName) =>
            self.NameMappings.GetValueOrDefault($"{contextName}.{sourceName}")
                ?? self.GeneratedName(sourceName);

        /// <summary>
        /// ブックスコープ定義名に対応する生成プロパティ名を決定します。
        /// </summary>
        internal string BookDefinedName(DefinedName definedName) =>
            self.GeneratedName("book", definedName.Name);

        /// <summary>
        /// ワークシートスコープ定義名に対応する生成プロパティ名を決定します。
        /// </summary>
        internal string SheetDefinedName(Worksheet sheet, DefinedName definedName) =>
            self.GeneratedName(sheet.Name, definedName.Name);

        /// <summary>
        /// Excelテーブル列に対応する生成プロパティ名を決定します。
        /// </summary>
        internal string TableColumn(Table table, TableColumn column) =>
            self.GeneratedName(table.Name, column.Name);
    }
}
