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
        var mapped = Activator.CreateInstance<T>();

        foreach (var property in MappedProperties)
        {
            property.SetValue(
                mapped,
                ConvertValue(GetSourceValue(row, property), property.PropertyType));
        }

        return mapped;
    }

    static object GetSourceValue(TableRow row, PropertyInfo property) =>
        row[GetColumnName(property)].Value;

    static string GetColumnName(PropertyInfo property) =>
        property.GetCustomAttribute<SpreadsheetColumnAttribute>()?.Name
            ?? throw new NotImplementedException();

    static IEnumerable<PropertyInfo> MappedProperties =>
        from property in typeof(T).GetProperties()
        where property.GetCustomAttribute<SpreadsheetColumnAttribute>() != null
        select property;

    static object ConvertValue(object sourceValue, Type propertyType) =>
        propertyType switch
        {
            _ when propertyType == typeof(int) => ConvertInteger(sourceValue),
            _ when propertyType == typeof(double) => ConvertDouble(sourceValue),
            _ when propertyType == typeof(string) => ConvertString(sourceValue),
            _ => throw new NotImplementedException()
        };

    static int ConvertInteger(object sourceValue) =>
        sourceValue is double number && double.IsInteger(number)
            ? (int)number
            : throw new NotImplementedException();

    static double ConvertDouble(object sourceValue) =>
        sourceValue is double number
            ? number
            : throw new NotImplementedException();

    static string ConvertString(object sourceValue) =>
        sourceValue is string text
            ? text
            : throw new NotImplementedException();
}
