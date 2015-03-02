using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public class CellRangeCollection
    {
        Dictionary<Tuple<string, string>, CellRange> cache = new Dictionary<Tuple<string, string>, CellRange>();

        public CellRange this[string topLeft, string bottomRight]
        {
            get
            {
                if(!cache.ContainsKey(Tuple.Create(topLeft, bottomRight)))
                {
                    cache[Tuple.Create(topLeft, bottomRight)] = new CellRange(topLeft, bottomRight);
                }

                return cache[Tuple.Create(topLeft, bottomRight)];
            }
        }
    }
}
