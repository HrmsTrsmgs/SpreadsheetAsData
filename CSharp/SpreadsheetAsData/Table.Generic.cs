using System.Reflection;

namespace Marimo.SpreadsheetAsData;

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
        ValidateAttributedProperties(HasPublicGetter);

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
    void Replace(TableRow row, T item)
    {
        foreach (var (cell, value) in
            from property in MappedProperties
            let sourceValue = property.GetValue(item)
            let cell = row[GetColumnName(property)]
            select (cell, Value: ConvertValueToCellValue(sourceValue, row, property)))
        {
            cell.Value = value;
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

        foreach (var (property, value) in
            from property in MappedProperties
            select (property, Value: ConvertValue(GetSourceValue(row, property), row, property)))
        {
            property.SetValue(mapped, value);
        }

        return mapped;
    }

    /// <summary>
    /// マッピング定義が、この Excel テーブルの列構成に対応できることを検証します。
    /// </summary>
    void ValidateColumns()
    {
        ValidateDuplicateColumns();
        ValidateAttributedProperties(HasPublicSetter);

        if ((
            from property in MappedProperties
            let columnName = GetColumnName(property)
            where !Columns.Contains(columnName)
            select (Property: property, ColumnName: columnName)
        ).TryGetFirst(out var missingColumn))
        {
            throw new TableMappingException(
                this,
                typeof(T),
                missingColumn.ColumnName,
                missingColumn.Property);
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
    /// 属性で対応付けたプロパティが、読み取りまたは書き込みに必要なアクセサーを持つか検証します。
    /// </summary>
    void ValidateAttributedProperties(Func<PropertyInfo, bool> hasPublicAccessor)
    {
        if ((
            from property in MappedProperties
            where HasSpreadsheetNameAttribute(property)
                && !hasPublicAccessor(property)
            select property
        ).TryGetFirst(out var propertyWithoutPublicAccessor))
        {
            throw new TableMappingException(
                this,
                typeof(T),
                GetColumnName(propertyWithoutPublicAccessor),
                propertyWithoutPublicAccessor);
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
        property.GetCustomAttribute<SpreadsheetNameAttribute>()?.Name
            ?? property.Name;

    /// <summary>
    /// プロパティに Excel テーブル列名を明示する属性があるかどうかを返します。
    /// </summary>
    /// <param name="property">確認するプロパティ。</param>
    /// <returns>列名を明示する属性がある場合は true。</returns>
    static bool HasSpreadsheetNameAttribute(PropertyInfo property) =>
        property.GetCustomAttribute<SpreadsheetNameAttribute>() is not null;

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
        HasSpreadsheetNameAttribute(property)
            || HasPublicGetter(property) && HasPublicSetter(property);

    /// <summary>
    /// 行のインスタンスに属する公開プロパティからマッピング対象を取得します。
    /// </summary>
    static IEnumerable<PropertyInfo> MappedProperties =>
        from property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
        where IsMappedProperty(property)
        select property;

    /// <summary>
    /// 元セル値を、指定したプロパティへ設定できる値へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のセル値。</param>
    /// <param name="row">変換元のテーブル行。</param>
    /// <param name="property">変換先プロパティ。</param>
    /// <returns>プロパティへ設定する値。</returns>
    object? ConvertValue(
        object sourceValue,
        TableRow row,
        PropertyInfo property) =>
        TableValueConversion.TryRead(sourceValue, property.PropertyType, out var converted)
            ? converted
            : throw ValueConversionException(sourceValue, row, property);

    /// <summary>
    /// マッピング元のプロパティ値を、セルへ設定できる値へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のプロパティ値。</param>
    /// <param name="row">書き込み先のテーブル行。</param>
    /// <param name="property">変換元プロパティ。</param>
    /// <returns>セルへ設定する値。</returns>
    object? ConvertValueToCellValue(
        object? sourceValue,
        TableRow row,
        PropertyInfo property) =>
        TableValueConversion.TryWrite(sourceValue, out var converted)
            ? converted
            : throw ValueConversionException(sourceValue, row, property);

    /// <summary>
    /// 読み書きどちらの変換失敗にも、同じ列・プロパティ・行の情報を付けます。
    /// </summary>
    TableMappingException ValueConversionException(object? sourceValue, TableRow row, PropertyInfo property) =>
        new(
            this,
            typeof(T),
            GetColumnName(property),
            property,
            row,
            sourceValue ?? new BlankValue());

}
