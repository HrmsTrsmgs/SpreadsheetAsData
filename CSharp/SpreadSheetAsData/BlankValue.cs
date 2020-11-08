using DocumentFormat.OpenXml.Math;

namespace Marimo.SpreadSheetAsData
{
    public struct BlankValue
    {
        public override bool Equals(object obj)
        {
            return obj is int && (int)obj == 0;
        }

        public static bool operator ==(BlankValue a, int b) => a.Equals(b);

        public static bool operator !=(BlankValue a, int b) => !(a == b);
    }
}