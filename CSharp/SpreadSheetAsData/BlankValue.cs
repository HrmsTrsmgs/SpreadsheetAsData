using DocumentFormat.OpenXml.Math;

namespace Marimo.SpreadSheetAsData
{
    public struct BlankValue
    {
        public override bool Equals(object obj) =>
            obj switch
            {
                int i => this == i,
                double d => this == d,
                string s => this == s,
                _ => false
            };
        public static implicit operator string(BlankValue _) => "";

        public static implicit operator double(BlankValue _) => .0;
    }
}