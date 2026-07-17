namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブル行を利用者定義型へ対応付ける処理に失敗したことを表します。
/// </summary>
public sealed class TableMappingException : Exception
{
    /// <summary>
    /// 空のマッピング例外を作成します。
    /// </summary>
    public TableMappingException()
    {
    }

    internal TableMappingException(
        Table table,
        Type mappingType,
        string columnName)
    {
        TableName = table.Name;
        MappingType = mappingType;
        ColumnName = columnName;
    }

    internal TableMappingException(
        Table table,
        Type mappingType,
        string columnName,
        System.Reflection.PropertyInfo property)
        : this(table, mappingType, columnName)
    {
        PropertyName = property.Name;
        PropertyType = property.PropertyType;
    }

    internal TableMappingException(
        Table table,
        Type mappingType,
        string columnName,
        System.Reflection.PropertyInfo property,
        TableRow row,
        object sourceValue)
        : this(table, mappingType, columnName, property)
    {
        WorksheetRowIndex = row.WorksheetRowIndex;
        SourceValue = sourceValue;
    }

    /// <summary>
    /// マッピングに失敗した Excel テーブル名を取得または設定します。
    /// </summary>
    public string? TableName { get; init; }

    /// <summary>
    /// マッピング先の型を取得または設定します。
    /// </summary>
    public Type? MappingType { get; init; }

    /// <summary>
    /// マッピングに失敗した Excel テーブル列名を取得または設定します。
    /// </summary>
    public string? ColumnName { get; init; }

    /// <summary>
    /// マッピングに失敗したプロパティ名を取得または設定します。
    /// </summary>
    public string? PropertyName { get; init; }

    /// <summary>
    /// マッピングに失敗したプロパティの型を取得または設定します。
    /// </summary>
    public Type? PropertyType { get; init; }

    /// <summary>
    /// 値変換に失敗したワークシート上の 1 始まりの行番号を取得または設定します。
    /// </summary>
    public uint? WorksheetRowIndex { get; init; }

    /// <summary>
    /// 値変換に失敗した元のセル値を取得または設定します。
    /// </summary>
    public object? SourceValue { get; init; }
}
