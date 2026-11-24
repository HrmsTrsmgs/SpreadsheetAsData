namespace Marimo.SpreadsheetAsData;

/// <summary>
/// XMLの格納形式に依存しないセル値を表します。空白は <see cref="BlankValue"/> で扱います。
/// </summary>
public readonly union TypedValue(string, double, bool, BlankValue);
