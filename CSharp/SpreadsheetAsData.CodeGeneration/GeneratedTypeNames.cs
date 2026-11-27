using Marimo.SpreadsheetAsData;

namespace Marimo.SpreadsheetAsData.CodeGeneration;

/// <summary>
/// 生成する型の名前、値から推論する型、生成名に隠されない参照型の表記を扱います。
/// </summary>
static class GeneratedTypeNames
{
    /// <summary>
    /// ファイル名から、ブックを型付きで表すクラスの識別子を決定します。
    /// </summary>
    internal static string BookTypeName(string filePath) =>
        $"{WorkbookFileIdentifier(filePath)}Book";

    /// <summary>
    /// 同じファイル名から、ブック全体のデータを表すクラスの識別子を決定します。
    /// </summary>
    internal static string DataTypeName(string filePath) =>
        $"{WorkbookFileIdentifier(filePath)}Data";

    static string WorkbookFileIdentifier(string filePath) =>
        Path.GetFileNameWithoutExtension(filePath).ToCSharpIdentifier();

    /// <summary>
    /// ワークシートの設定済み生成名を、派生型の識別子に変換します。
    /// </summary>
    internal static string SheetTypeName(Worksheet sheet, CodeGenerationOptions options) =>
        $"{options.GeneratedName(sheet.Name)}Sheet";

    /// <summary>
    /// 対応する行データ型の識別子からテーブルの派生型名を決定します。
    /// </summary>
    internal static string TableTypeName(string rowTypeName) =>
        $"{rowTypeName}Table";

    /// <summary>
    /// 既存の型付きTableマッピングで読み込めるプロパティ型名を、列の値から決定します。
    /// </summary>
    internal static string ColumnPropertyTypeName(Table table, TableColumn column)
    {
        var values = (
            from row in table.Rows
            select row[column].TypedValue
        ).ToArray();

        if (values.All(it => it is string or BlankValue))
        {
            return "string";
        }

        var nonBlankValues =
            from value in values
            where value is not BlankValue
            select value;
        var typeName =
            nonBlankValues.All(CanConvertToInt32)
                ? "int"
            : nonBlankValues.All(it => it is double)
                ? "double"
            : nonBlankValues.All(it => it is bool)
                ? "bool"
            : "dynamic";

        return typeName != "dynamic" && values.Any(it => it is BlankValue)
            ? $"{typeName}?"
            : typeName;

        // セル値がintの範囲に収まる整数値かどうかを判定します。
        static bool CanConvertToInt32(TypedValue value) =>
            value is double number
            && TableValueConversion.CanConvertToInt32(number);
    }

    /// <summary>
    /// セルの現在値を表すC#プロパティ型名を返します。
    /// </summary>
    internal static string CellValueTypeName(TypedValue value) =>
        value switch
        {
            string => "string",
            double => "double",
            bool => "bool",
            _ => "dynamic"
        };

    /// <summary>
    /// 行データ型が参照先の型名を隠す場合だけ完全修飾し、それ以外は短い型名を返します。
    /// 式中ではBookのプロパティも型名を隠すため、シート名と定義名を確認します。
    /// その他の生成型には接尾辞が付くため、ここで扱う既存型名とは衝突しません。
    /// </summary>
    internal static string ReferencedTypeName(
        string typeName,
        Workbook book,
        CodeGenerationOptions options,
        bool usedAsExpression = false)
    {
        var shortName = typeName.Split('.')[^1];
        var isHidden = book.Tables.Any(table => options.GeneratedName(table.Name).IdentifierValue == shortName)
            || usedAsExpression && (
                book.Sheets.Values.Any(sheet => options.GeneratedName(sheet.Name).IdentifierValue == shortName)
                || book.DefinedNames.Any(name => name.Worksheet is null
                    && options.BookDefinedName(name).IdentifierValue == shortName));
        return !isHidden
            ? shortName
        : typeName.Contains('.')
            ? $"global::{typeName}"
        : $"global::Marimo.SpreadsheetAsData.{typeName}";
    }

    /// <summary>
    /// 行データ型が属性クラス名を隠す場合だけ完全修飾し、それ以外は属性の短縮表記を使います。
    /// </summary>
    internal static string ReferencedAttributeName(Workbook book, CodeGenerationOptions options)
    {
        var typeName = ReferencedTypeName(nameof(SpreadsheetNameAttribute), book, options);
        return typeName == nameof(SpreadsheetNameAttribute) ? "SpreadsheetName" : typeName;
    }
}
