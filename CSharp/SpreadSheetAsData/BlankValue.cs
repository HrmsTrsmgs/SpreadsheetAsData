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
        

        public static bool operator ==(BlankValue a, int b) => b == 0;

        public static bool operator !=(BlankValue a, int b) => !(a == b);

        public static bool operator ==(BlankValue a, string b) => b == "";

        public static bool operator !=(BlankValue a, string b) => !(a == b);
        public static bool operator ==(string a, BlankValue b) => a == "";

        public static bool operator !=(string a, BlankValue b) => !(a == b);



        public static bool operator ==(BlankValue a, double b) => b == 0;

        public static bool operator !=(BlankValue a, double b) => !(a == b);
    }
}