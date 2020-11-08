using DocumentFormat.OpenXml.Math;

namespace Marimo.SpreadSheetAsData
{
    public struct BlankValue
    {
        public override bool Equals(object obj)
        {
            return obj is int && (int)obj == 0 || obj is string && (string)obj == "";
        }

        public static bool operator ==(BlankValue a, int b) => a.Equals(b);

        public static bool operator !=(BlankValue a, int b) => !(a == b);

        public static bool operator ==(BlankValue a, string b) => a.Equals(b);

        public static bool operator !=(BlankValue a, string b) => !(a == b);


        public static bool operator ==(BlankValue a, double b) => a.Equals(b);

        public static bool operator !=(BlankValue a, double b) => !(a == b);
    }
}