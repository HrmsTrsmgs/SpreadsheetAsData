using System.Reflection;

namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブック内の定義名とExcelテーブルを、データオブジェクトのプロパティへ対応付けます。
/// </summary>
sealed class WorkbookDataMapper(Workbook book)
{
    internal T Read<T>()
    {
        // 値型でも各SetValueが同じインスタンスを更新するよう、一度だけボックス化します。
        object? data = Activator.CreateInstance<T>();

        foreach (var property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            property.SetValue(data, ReadPropertyValue(property));
        }

        // Tとして作った値を戻します。nullの場合も元のTが許すnullのまま返します。
        return (T)data!;

        // プロパティに対応するテーブル、単一セル、範囲から設定する値を読み取ります。
        object? ReadPropertyValue(PropertyInfo property)
        {
            if (TryGetTable(property, out var table))
            {
                return CreateTypedTable(table, TableRowType(property));
            }

            var range = DefinedNameRange(property);

            if (range.TopLeftCell == range.BottomRightCell)
            {
                return ConvertValue(range.TopLeftCell.Value, property.PropertyType);
            }

            return range.Values;
        }

        // 空白値をstringプロパティへ設定できる値に変換します。
        static object ConvertValue(object value, Type propertyType) =>
            (value, propertyType) switch
            {
                (BlankValue, var type) when type == typeof(string) => "",
                _ => value
            };
    }

    internal void Replace<T>(T data)
    {
        foreach (var property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (TryGetTable(property, out var table))
            {
                ReplaceTableRows(
                    table,
                    TableRowType(property),
                    property.GetValue(data));
                continue;
            }

            var range = DefinedNameRange(property);

            if (range.TopLeftCell == range.BottomRightCell)
            {
                range.TopLeftCell.Value = property.GetValue(data);
                continue;
            }

            range.Values =
                (IEnumerable<IEnumerable<object?>>)(
                    property.GetValue(data)
                        ?? Array.Empty<IEnumerable<object?>>());
        }
    }

    /// <summary>
    /// 属性で指定した名前を優先し、対応するExcelテーブルを検索します。
    /// </summary>
    bool TryGetTable(PropertyInfo property, out Table table)
    {
        var tableName = property
            .GetCustomAttribute<SpreadSheetNameAttribute>()
            ?.Name;

        return
            (
                from candidate in book.Tables
                where tableName is not null
                    ? candidate.Name == tableName
                    : candidate.Name.ToCSharpIdentifier() == property.Name
                select candidate
            ).TryGetFirst(out table);
    }

    /// <summary>
    /// Excelテーブルへ対応付けるプロパティから行データ型を取得します。
    /// </summary>
    static Type TableRowType(PropertyInfo property) =>
        property.PropertyType.GenericTypeArguments.Single();

    /// <summary>
    /// 実行時に決まる行データ型を使用して型付きテーブルを作成します。
    /// </summary>
    static object? CreateTypedTable(Table table, Type rowType) =>
        (
            from method in typeof(Table).GetMethods()
            where method.Name == nameof(Table.Enumerate)
            where method.IsGenericMethodDefinition
            select method
        ).Single()
        .MakeGenericMethod(rowType)
        .Invoke(table, null);

    /// <summary>
    /// 実行時に決まる行データ型に対応した置換処理を呼び出します。
    /// </summary>
    static void ReplaceTableRows(Table table, Type rowType, object? rows) =>
        (
            from method in typeof(Table<>).MakeGenericType(rowType).GetMethods()
            where method.Name == nameof(Table<>.Replace)
            select method
        ).Single()
        .Invoke(
            CreateTypedTable(table, rowType),
            [rows]);

    CellRange DefinedNameRange(PropertyInfo property)
    {
        var attribute =
            property.GetCustomAttribute<SpreadSheetNameAttribute>();

        if (attribute is not null)
        {
            return attribute.WorksheetName is string worksheetName
                ? book.Sheets[worksheetName].Range[attribute.Name]
                : book.Range[attribute.Name];
        }

        return
            (
                from definedName in book.DefinedNames
                where definedName.Name.ToCSharpIdentifier() == property.Name
                select definedName.Range
            ).Single();
    }
}
