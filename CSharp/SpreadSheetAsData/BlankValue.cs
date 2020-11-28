using DocumentFormat.OpenXml.Math;

namespace Marimo.SpreadSheetAsData
{
    public struct BlankValue
    { 
        public static implicit operator string(BlankValue _) => "";

        public static implicit operator double(BlankValue _) => .0;

        public override string ToString() => "{Blank}";
    }
}