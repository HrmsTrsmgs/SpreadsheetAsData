using System.Reflection;

namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブルの各データ行を指定した型へ対応付けて列挙する型付きテーブルを表します。
/// </summary>
/// <typeparam name="T">各データ行を対応付ける型。</typeparam>
public sealed class Table<T> : Table, IEnumerable<T>
{
    /// <summary>
    /// 指定した非型付き Excel テーブルから型付きテーブルを作成します。
    /// </summary>
    /// <param name="source">型付き列挙の元になる Excel テーブル。</param>
    internal Table(Table source)
        : base(source.TableDefinitionPart, source.Worksheet)
    {
        ValidateColumns();
    }

    /// <summary>
    /// Excel テーブルの各データ行を <typeparamref name="T"/> へ対応付けて列挙します。
    /// </summary>
    /// <returns>型付き行の列挙子。</returns>
    public IEnumerator<T> GetEnumerator() =>
        (
            from row in base.Rows
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
        ValidateDuplicateColumns();

        foreach (var property in MappedProperties)
        {
            var columnName = GetColumnName(property);

            if (!Columns.Contains(columnName))
            {
                throw new TableMappingException(
                    this,
                    typeof(T),
                    columnName,
                    property);
            }
        }
    }

    void ValidateDuplicateColumns()
    {
        var duplicateColumnName = FindDuplicateColumnName();

        if (duplicateColumnName is not null)
        {
            throw new TableMappingException(
                this,
                typeof(T),
                duplicateColumnName);
        }
    }

    static string? FindDuplicateColumnName() =>
        (
            from property in MappedProperties
            group property by GetColumnName(property) into propertiesByColumn
            where propertiesByColumn.Skip(1).Any()
            select propertiesByColumn.Key
        ).FirstOrDefault();

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

        throw new TableMappingException(
            this,
            typeof(T),
            GetColumnName(property),
            property,
            row,
            sourceValue);
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
