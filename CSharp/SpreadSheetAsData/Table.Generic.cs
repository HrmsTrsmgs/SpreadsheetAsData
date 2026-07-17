using System.Reflection;

namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブルの各データ行を指定した型へ対応付けて列挙する型付きテーブルを表します。
/// </summary>
/// <typeparam name="T">各データ行を対応付ける型。</typeparam>
public sealed class Table<T> : IEnumerable<T>
{
    /// <summary>
    /// 型付き列挙の元になる非型付き Excel テーブルです。
    /// </summary>
    readonly Table source;

    /// <summary>
    /// 指定した非型付き Excel テーブルから型付きテーブルを作成します。
    /// </summary>
    /// <param name="source">型付き列挙の元になる Excel テーブル。</param>
    internal Table(Table source)
    {
        this.source = source;
    }

    /// <summary>
    /// Excel テーブルの各データ行を <typeparamref name="T"/> へ対応付けて列挙します。
    /// </summary>
    /// <returns>型付き行の列挙子。</returns>
    public IEnumerator<T> GetEnumerator() =>
        (
            from row in source.Rows
            select Map(row)
        ).GetEnumerator();

    /// <summary>
    /// Excel テーブルの各データ行を <typeparamref name="T"/> へ対応付けて列挙します。
    /// </summary>
    /// <returns>型付き行の列挙子。</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        GetEnumerator();

    T Map(TableRow row)
    {
        ValidateColumns();

        var mapped = Activator.CreateInstance<T>();

        foreach (var property in MappedProperties)
        {
            var sourceValue = GetSourceValue(row, property);

            property.SetValue(
                mapped,
                ConvertValue(sourceValue, row, property));
        }

        return mapped;
    }

    void ValidateColumns()
    {
        foreach (var property in MappedProperties)
        {
            var columnName = GetColumnName(property);

            if (!source.Columns.Contains(columnName))
            {
                throw CreateMappingException(columnName, property);
            }
        }
    }

    TableMappingException CreateMappingException(
        string columnName,
        PropertyInfo property) =>
        new()
        {
            TableName = source.Name,
            MappingType = typeof(T),
            ColumnName = columnName,
            PropertyName = property.Name,
            PropertyType = property.PropertyType
        };

    TableMappingException CreateMappingException(
        object sourceValue,
        TableRow row,
        PropertyInfo property) =>
        new()
        {
            TableName = source.Name,
            MappingType = typeof(T),
            ColumnName = GetColumnName(property),
            PropertyName = property.Name,
            PropertyType = property.PropertyType,
            WorksheetRowIndex = row.WorksheetRowIndex,
            SourceValue = sourceValue
        };

    static object GetSourceValue(TableRow row, PropertyInfo property) =>
        row[GetColumnName(property)].Value;

    static string GetColumnName(PropertyInfo property) =>
        property.GetCustomAttribute<SpreadsheetColumnAttribute>()?.Name
            ?? property.Name;

    static IEnumerable<PropertyInfo> MappedProperties =>
        from property in typeof(T).GetProperties()
        select property;

    object ConvertValue(
        object sourceValue,
        TableRow row,
        PropertyInfo property)
    {
        if (TryConvertValue(sourceValue, property.PropertyType, out var converted))
        {
            return converted;
        }

        throw CreateMappingException(sourceValue, row, property);
    }

    static bool TryConvertValue(
        object sourceValue,
        Type propertyType,
        out object converted)
    {
        object? conversion = (propertyType, sourceValue) switch
        {
            ({ } type, double number) when type == typeof(int)
                && double.IsInteger(number) => (int)number,
            ({ } type, double number) when type == typeof(double) => number,
            ({ } type, string text) when type == typeof(string) => text,
            _ => null
        };

        converted = conversion ?? new();

        return conversion != null;
    }
}
