using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public class CellRange
    {
        private string bottomRight;
        private string topLeft;

        public CellRange(string topLeft, string bottomRight)
        {
            this.topLeft = topLeft;
            this.bottomRight = bottomRight;
        }

        public override string ToString()
        {
            return topLeft + ":" + bottomRight;
        }
    }
}
