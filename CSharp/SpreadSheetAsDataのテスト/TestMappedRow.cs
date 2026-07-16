using Marimo.SpreadSheetAsData;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class TestMappedRow
{
    [SpreadsheetColumn("int")]
    public int IntegerValue { get; set; }

    [SpreadsheetColumn("float")]
    public double FloatingPointValue { get; set; }

    [SpreadsheetColumn("string")]
    public string TextValue { get; set; } = "";
}
