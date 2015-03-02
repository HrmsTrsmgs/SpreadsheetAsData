using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public class CellRangeCollection
    {
        public CellRange this[string topLeft, string bottomRight]
        {
            get
            {
                return new CellRange(topLeft, bottomRight);
            }
        }
    }
}
