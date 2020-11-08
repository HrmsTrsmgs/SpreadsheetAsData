using System.Collections.Generic;

namespace Marimo.SpreadSheetAsData
{
    public class CellRangeCollection
    {
        Dictionary<(string, string), CellRange> cache = new Dictionary<(string, string), CellRange>();

        public CellRange this[string topLeft, string bottomRight]
        {
            get
            {
                if(!cache.ContainsKey((topLeft, bottomRight)))
                {
                    cache[(topLeft, bottomRight)] = new CellRange(topLeft, bottomRight);
                }

                return cache[(topLeft, bottomRight)];
            }
        }
    }
}
