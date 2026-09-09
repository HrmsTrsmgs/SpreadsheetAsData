using Marimo.SpreadSheetAsData;

namespace Marimo.SpreadSheetAsData.Test;

public sealed class TestMappedRow
{
    [SpreadSheetName("数値2")]
    public int IntegerValue { get; set; }

    [SpreadSheetName("数値")]
    public double FloatingPointValue { get; set; }

    [SpreadSheetName("文字列")]
    public string TextValue { get; set; } = "";
}
