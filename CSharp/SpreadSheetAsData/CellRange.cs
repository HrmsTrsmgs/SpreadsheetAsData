using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public class CellRange
    {
        private readonly string bottomRight;
        private readonly string topLeft;

        public CellRange(string topLeft, string bottomRight)
        {
            this.topLeft = topLeft;
            this.bottomRight = bottomRight;
        }

        public override string ToString() => topLeft + ":" + bottomRight;
    }
}
