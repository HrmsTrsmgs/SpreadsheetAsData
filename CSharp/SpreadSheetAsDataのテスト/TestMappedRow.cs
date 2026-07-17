using Marimo.SpreadSheetAsData;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class TestMappedRow
{
    [SpreadsheetColumn("数値2")]
    public int IntegerValue { get; set; }

    [SpreadsheetColumn("数値")]
    public double FloatingPointValue { get; set; }

    [SpreadsheetColumn("文字列")]
    public string TextValue { get; set; } = "";
}
