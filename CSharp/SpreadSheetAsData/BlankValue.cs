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
        

        public static bool operator ==(BlankValue left, int right) => right == 0;

        public static bool operator !=(BlankValue left, int right) => !(left == right);

        public static bool operator ==(int left, BlankValue right) => right == 0;

        public static bool operator !=(int left, BlankValue right) => !(left == right);

        public static bool operator ==(BlankValue left, string right) => right == "";

        public static bool operator !=(BlankValue left, string right) => !(left == right);
        public static bool operator ==(string left, BlankValue right) => left == "";

        public static bool operator !=(string left, BlankValue right) => !(left == right);

        public static bool operator ==(BlankValue left, double right) => right == 0;

        public static bool operator !=(BlankValue left, double right) => !(left == right);

        public static bool operator ==(double left, BlankValue right) => left == 0;

        public static bool operator !=(double left, BlankValue right) => !(left == right);
    }
}