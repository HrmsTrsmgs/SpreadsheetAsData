using System.Reflection;

namespace Marimo.SpreadsheetAsData;

/// <summary>
/// Excel テーブルの各データ行を指定した型へ対応付けて列挙する型付きテーブルを表します。
/// </summary>
/// <typeparam name="T">各データ行を対応付ける型。</typeparam>
public class Table<T> : Table, IEnumerable<T>
{
    readonly record struct ColumnMapping(PropertyInfo Property, string? ExplicitName)
    {
        internal string ColumnName => ExplicitName ?? Property.Name;
    }

    /// <summary>
    /// 検証と各行の読み書きで共有する列対応です。セルやプロパティ値は保持しません。
    /// </summary>
    readonly ColumnMapping[] mappings;

    /// <summary>
    /// 指定した非型付き Excel テーブルから型付きテーブルを作成します。
    /// </summary>
    /// <param name="source">型付き列挙の元になる Excel テーブル。</param>
    protected internal Table(Table source)
        : base(source.TableDefinitionPart, source.Worksheet)
    {
        ValidateMappingTypeCanBeCreated();
        mappings = [
            .. from property in typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
               let name = property.GetCustomAttribute<SpreadsheetNameAttribute>()?.Name
               where name is not null || HasPublicGetter(property) && HasPublicSetter(property)
               select new ColumnMapping(property, name)
        ];
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
            from mapping in mappings
            let sourceValue = mapping.Property.GetValue(item)
            let cell = row[mapping.ColumnName]
            select (cell, Value: ConvertValueToCellValue(sourceValue, row, mapping)))
        {
            cell.TypedValue = value;
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
        object mapped = Activator.CreateInstance<T>() ?? throw new TableMappingException(this, typeof(T));

        foreach (var (property, value) in
            from mapping in mappings
            let cell = row[mapping.ColumnName]
            select (mapping.Property, Value: ConvertValue(cell.TypedValue, row, mapping)))
        {
            property.SetValue(mapped, value);
        }

        return (T)mapped;
    }

    /// <summary>
    /// マッピング定義が、この Excel テーブルの列構成に対応できることを検証します。
    /// </summary>
    void ValidateColumns()
    {
        ValidateDuplicateColumns();
        ValidateAttributedProperties(HasPublicSetter);

        if (mappings.Where(it => !Columns.Contains(it.ColumnName)).TryGetFirst(out var missingColumn))
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
            from mapping in mappings
            group mapping by mapping.ColumnName into propertiesByColumn
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
            from mapping in mappings
            where mapping.ExplicitName is not null && !hasPublicAccessor(mapping.Property)
            select mapping
        ).TryGetFirst(out var invalidMapping))
        {
            throw new TableMappingException(
                this,
                typeof(T),
                invalidMapping.ColumnName,
                invalidMapping.Property);
        }
    }

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
    /// 元セル値を、指定したプロパティへ設定できる値へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のセル値。</param>
    /// <param name="row">変換元のテーブル行。</param>
    /// <param name="mapping">変換先プロパティと列の対応。</param>
    /// <returns>プロパティへ設定する値。</returns>
    object? ConvertValue(
        TypedValue sourceValue,
        TableRow row,
        ColumnMapping mapping) =>
        TableValueConversion.TryRead(sourceValue, mapping.Property.PropertyType, out var converted)
            ? converted
            : throw ValueConversionException(sourceValue.Value, row, mapping);

    /// <summary>
    /// マッピング元のプロパティ値を、セルへ設定できる値へ変換します。
    /// </summary>
    /// <param name="sourceValue">変換元のプロパティ値。</param>
    /// <param name="row">書き込み先のテーブル行。</param>
    /// <param name="mapping">変換元プロパティと列の対応。</param>
    /// <returns>セルへ設定する値。</returns>
    TypedValue ConvertValueToCellValue(
        object? sourceValue,
        TableRow row,
        ColumnMapping mapping) =>
        TableValueConversion.TryWrite(sourceValue, out var converted)
            ? converted
            : throw ValueConversionException(sourceValue, row, mapping);

    /// <summary>
    /// 読み書きどちらの変換失敗にも、同じ列・プロパティ・行の情報を付けます。
    /// </summary>
    TableMappingException ValueConversionException(object? sourceValue, TableRow row, ColumnMapping mapping) =>
        new(
            this,
            typeof(T),
            mapping.ColumnName,
            mapping.Property,
            row,
            sourceValue ?? new BlankValue());
}
