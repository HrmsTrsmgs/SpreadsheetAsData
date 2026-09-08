using System.Reflection;

namespace Marimo.SpreadSheetAsData;

/// <summary>
/// ブック内の定義名とExcelテーブルを、データオブジェクトのプロパティへ対応付けます。
/// </summary>
sealed class WorkbookDataMapper(Workbook book)
{
    internal T Read<T>()
    {
        var data = Activator.CreateInstance<T>();

        foreach (var property in typeof(T).GetProperties())
        {
            if (TryGetTable(property, out var table))
            {
                property.SetValue(
                    data,
                    CreateTypedTable(table, TableRowType(property)));
                continue;
            }

            var range = DefinedNameRange(property);

            property.SetValue(
                data,
                range.TopLeftCell == range.BottomRightCell
                    ? range.TopLeftCell.Value
                    : range.Values);
        }

        return data;
    }

    internal void Replace<T>(T data)
    {
        foreach (var property in typeof(T).GetProperties())
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
    /// プロパティ名と同じC#識別子になるExcelテーブルを検索します。
    /// </summary>
    bool TryGetTable(PropertyInfo property, out Table table) =>
        (
            from candidate in book.Tables
            where candidate.Name.ToCSharpIdentifier() == property.Name
            select candidate
        ).TryGetFirst(out table);

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
            where method.Name == nameof(Table<object>.Replace)
            select method
        ).Single()
        .Invoke(
            CreateTypedTable(table, rowType),
            [rows]);

    CellRange DefinedNameRange(PropertyInfo property)
    {
        var attribute =
            property.GetCustomAttribute<SpreadsheetDefinedNameAttribute>();

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
