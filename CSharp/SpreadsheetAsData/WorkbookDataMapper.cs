using System.Reflection;

namespace Marimo.SpreadsheetAsData;

/// <summary>
/// ブック内の定義名とExcelテーブルを、データオブジェクトのプロパティへ対応付けます。
/// </summary>
sealed class WorkbookDataMapper(Workbook book)
{
    readonly union PropertyTarget(Table, Cell, CellRange);

    internal T Read<T>()
    {
        // 値型でも各SetValueが同じインスタンスを更新するよう、一度だけボックス化します。
        object? data = Activator.CreateInstance<T>();

        foreach (var (property, value) in
            from property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            select (property, Value: ReadPropertyValue(property)))
        {
            property.SetValue(data, value);
        }

        // Tとして作った値を戻します。nullの場合も元のTが許すnullのまま返します。
        return (T)data!;
    }

    object? ReadPropertyValue(PropertyInfo property) =>
        ResolveTarget(property) switch
        {
            Table table => CreateTypedTable(table, TableRowType(property)),
            Cell cell => ReadCellValue(cell.TypedValue, property.PropertyType),
            CellRange range => range.Values
        };

    static object? ReadCellValue(TypedValue value, Type propertyType) =>
        value is BlankValue && propertyType == typeof(string) ? "" : value.Value;

    internal void Replace<T>(T data)
    {
        foreach (var (property, target) in
            from property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
            where property.CanRead
            select (property, Target: ResolveTarget(property)))
        {
            ReplacePropertyValue(property, target, data);
        }
    }

    static void ReplacePropertyValue<T>(PropertyInfo property, PropertyTarget target, T data)
    {
        switch (target)
        {
            case Table table:
                ReplaceTableRows(table, TableRowType(property), property.GetValue(data));
                break;
            case Cell cell:
                cell.TypedValue = Cell.NormalizeValue(property.GetValue(data));
                break;
            case CellRange range:
                range.Values = (IEnumerable<IEnumerable<object?>>)(
                    property.GetValue(data) ?? Array.Empty<IEnumerable<object?>>());
                break;
        }
    }

    /// <summary>
    /// テーブルを定義名より優先します。二次元の値を受け取る場合は一セルでも範囲として扱います。
    /// </summary>
    PropertyTarget ResolveTarget(PropertyInfo property)
    {
        var attribute = property.GetCustomAttribute<SpreadsheetNameAttribute>();
        if (FindTable(property, attribute?.Name) is { } table)
        {
            return table;
        }

        var range = DefinedNameRange(property, attribute);
        var topLeft = range.TopLeftCell;
        return topLeft == range.BottomRightCell
            && property.PropertyType != typeof(IEnumerable<IEnumerable<object?>>)
            ? topLeft
            : range;
    }

    /// <summary>
    /// 属性で指定した名前を優先し、対応するExcelテーブルを検索します。
    /// </summary>
    Table? FindTable(PropertyInfo property, string? tableName) =>
        (
            from candidate in book.Tables
            where tableName is not null
                ? candidate.Name == tableName
                : candidate.Name.ToCSharpIdentifier() == property.Name
            select candidate
        ).FirstOrDefault();

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

    CellRange DefinedNameRange(PropertyInfo property, SpreadsheetNameAttribute? attribute)
    {
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
