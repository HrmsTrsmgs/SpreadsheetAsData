using Marimo.SpreadsheetAsData;

namespace Marimo.SpreadsheetAsData.Test;

public sealed class TestMappedRow
{
    [SpreadsheetName("数値2")]
    public int IntegerValue { get; set; }

    [SpreadsheetName("数値")]
    public double FloatingPointValue { get; set; }

    [SpreadsheetName("文字列")]
    public string TextValue { get; set; } = "";
}
