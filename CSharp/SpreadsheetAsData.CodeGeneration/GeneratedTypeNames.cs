using Marimo.SpreadsheetAsData;

namespace Marimo.SpreadsheetAsData.CodeGeneration;

/// <summary>
/// 値からの型推論と、生成名に隠されない型名の表記を扱います。宣言テンプレートの構文から分離します。
/// </summary>
static class GeneratedTypeNames
{
    /// <summary>
    /// 既存の型付きTableマッピングで読み込めるプロパティ型名を、列の値から決定します。
    /// </summary>
    internal static string ColumnPropertyTypeName(Table table, TableColumn column)
    {
        var values = (
            from row in table.Rows
            select row[column].Value
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
        static bool CanConvertToInt32(object value) =>
            value is double number
            && double.IsInteger(number)
            && number is >= int.MinValue and <= int.MaxValue;
    }

    /// <summary>
    /// セルの現在値を表すC#プロパティ型名を返します。
    /// </summary>
    internal static string CellValueTypeName(object value) =>
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
