using System.Collections.Generic;

namespace Marimo.SpreadSheetAsData
{
    public class CellRangeCollection
    {
        readonly Dictionary<(string TopLeft, string BottomRight), CellRange> cache = new();

        public CellRange this[string topLeft, string bottomRight]
        {
            get
            {
                var key = (topLeft, bottomRight);
                if (!cache.TryGetValue(key, out var range))
                {
                    range = new CellRange(topLeft, bottomRight);
                    cache[key] = range;
                }

                return range;
            }
        }
    }
}
