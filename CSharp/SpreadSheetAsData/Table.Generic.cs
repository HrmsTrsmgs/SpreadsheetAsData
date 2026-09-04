using System.Reflection;

namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブルの各データ行を指定した型へ対応付けて列挙する型付きテーブルを表します。
/// </summary>
/// <typeparam name="T">各データ行を対応付ける型。</typeparam>
public class Table<T> : Table, IEnumerable<T>
{
    /// <summary>
    /// 指定した非型付き Excel テーブルから型付きテーブルを作成します。
    /// </summary>
    /// <param name="source">型付き列挙の元になる Excel テーブル。</param>
    protected internal Table(Table source)
        : base(source.TableDefinitionPart, source.Worksheet)
    {
        ValidateMappingTypeCanBeCreated();
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

    /// <summary>
    /// Excel テーブルの内容を、指定した型付き行の内容で置き換えます。
    /// </summary>
    /// <param name="items">置き換え後の型付き行。</param>
    public void Replace(IEnumerable<T> items)
    {
        ValidateAttributedPropertiesHavePublicGetters();

        foreach (var (row, item) in base.Rows.Zip(items))
        {
            Replace(row, item);
        }
    }

    /// <summary>
    /// 非型付きのテーブル行へ、マッピング元の型付き行から値を書き込みます。
    /// </summary>
    /// <param name="row">書き込み先のテーブル行。</param>
    /// <param name="item">書き込み元の型付き行。</param>
    static void Replace(TableRow row, T item)
    {
        foreach (var property in MappedProperties)
        {
            row[GetColumnName(property)].Value = property.GetValue(item);
        }
    }

    /// <summary>
    /// マッピング先の型を作成できることを検証します。
    /// </summary>
    void ValidateMappingTypeCanBeCreated()
    {
        if (typeof(T).GetConstructor(Type.EmptyTypes) is null)
        {
            throw new TableMappingException(this, typeof(T));
        }
    }

    /// <summary>
    /// 非型付きのテーブル行を、マッピング先の型へ変換します。
    /// </summary>
    /// <param name="row">変換元のテーブル行。</param>
    /// <returns>変換した型付き行。</returns>
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

    /// <summary>
    /// マッピング定義が、この Excel テーブルの列構成に対応できることを検証します。
    /// </summary>
    void ValidateColumns()
    {
        ValidateDuplicateColumns();
        ValidateAttributedPropertiesHavePublicSetters();

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

    /// <summary>
    /// 同じ Excel テーブル列へ複数のプロパティを対応付けていないことを検証します。
    /// </summary>
    void ValidateDuplicateColumns()
    {
        if ((
            from property in MappedProperties
            group property by GetColumnName(property) into propertiesByColumn
            where propertiesByColumn.Skip(1).Any()
            select propertiesByColumn.Key
        ).TryGetFirst(out var duplicateColumnName))
        {
            throw new TableMappingException(
                this,
                typeof(T),
                duplicateColumnName);
        }
    }

    /// <summary>
    /// 属性で対応付けたプロパティへ値を書き込めることを検証します。
    /// </summary>
    void ValidateAttributedPropertiesHavePublicSetters()
    {
        if ((
            from property in MappedProperties
            where HasSpreadsheetColumnAttribute(property)
                && !HasPublicSetter(property)
            select property
        ).TryGetFirst(out var propertyWithoutPublicSetter))
        {
            throw new TableMappingException(
                this,
                typeof(T),
                GetColumnName(propertyWithoutPublicSetter),
                propertyWithoutPublicSetter);
        }
    }

    /// <summary>
    /// 属性で対応付けたプロパティから値を読み取れることを検証します。
    /// </summary>
    void ValidateAttributedPropertiesHavePublicGetters()
    {
        if ((
            from property in MappedProperties
            where HasSpreadsheetColumnAttribute(property)
                && !HasPublicGetter(property)
            select property
        ).TryGetFirst(out var propertyWithoutPublicGetter))
        {
            throw new TableMappingException(
                this,
                typeof(T),
                GetColumnName(propertyWithoutPublicGetter),
                propertyWithoutPublicGetter);
        }
    }

    /// <summary>
    /// テーブル行から、指定したプロパティに対応する元セル値を取得します。
    /// </summary>
    /// <param name="row">値を取得するテーブル行。</param>
    /// <param name="property">マッピング先プロパティ。</param>
    /// <returns>プロパティに対応するセル値。</returns>
    static object GetSourceValue(TableRow row, PropertyInfo property) =>
        row[GetColumnName(property)].Value;

    /// <summary>
    /// プロパティに対応する Excel テーブル列名を取得します。
    /// </summary>
    /// <param name="property">列名を取得するプロパティ。</param>
    /// <returns>属性で指定した列名。属性がない場合はプロパティ名。</returns>
    static string GetColumnName(PropertyInfo property) =>
        property.GetCustomAttribute<SpreadsheetColumnAttribute>()?.Name
            ?? property.Name;

    /// <summary>
    /// プロパティに Excel テーブル列名を明示する属性があるかどうかを返します。
    /// </summary>
    /// <param name="property">確認するプロパティ。</param>
    /// <returns>列名を明示する属性がある場合は true。</returns>
    static bool HasSpreadsheetColumnAttribute(PropertyInfo property) =>
        property.GetCustomAttribute<SpreadsheetColumnAttribute>() is not null;

    /// <summary>
    /// プロパティに public setter があるかどうかを返します。
    /// </summary>
    /// <param name="property">確認するプロパティ。</param>
    /// <returns>public setter がある場合は true。</returns>
    static bool HasPublicSetter(PropertyInfo property) =>
        property.SetMethod?.IsPublic == true;

    /// <summary>
    /// プロパティに public getter があるかどうかを返します。
    /// </summary>
    /// <param name="property">確認するプロパティ。</param>
    /// <returns>public getter がある場合は true。</returns>
    static bool HasPublicGetter(PropertyInfo property) =>
        property.GetMethod?.IsPublic == true;

    /// <summary>
    /// プロパティがマッピング対象かどうかを返します。
    /// </summary>
    /// <param name="property">確認するプロパティ。</param>
    /// <returns>列属性を持つ、または public getter と public setter を持つ場合は true。</returns>
    static bool IsMappedProperty(PropertyInfo property) =>
        HasSpreadsheetColumnAttribute(property)
            || HasPublicGetter(property) && HasPublicSetter(property);

    /// <summary>
    /// マッピング対象になる公開プロパティを取得します。
    /// </summary>
    static IEnumerable<PropertyInfo> MappedProperties =>
        from property in typeof(T).GetProperties()
        where IsMappedProperty(property)
        select property;

    /// <summary>
    /// 元セル値を、指定したプロパティへ設定できる値へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のセル値。</param>
    /// <param name="row">変換元のテーブル行。</param>
    /// <param name="property">変換先プロパティ。</param>
    /// <returns>プロパティへ設定する値。</returns>
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

    /// <summary>
    /// 元セル値を指定した型へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のセル値。</param>
    /// <param name="propertyType">変換先のプロパティ型。</param>
    /// <param name="converted">変換に成功した場合の値。</param>
    /// <returns>変換できた場合は true。</returns>
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
