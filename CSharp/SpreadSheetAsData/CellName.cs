using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public class CellName
    {
        string name;

        public CellName(string name)
        {
            this.name = name;
        }

        public uint ColumnIndex
        {
            get
            {
                return (uint)(ColumnName.First() - 'A') + 1;
            }
        }

        public string ColumnName
        {
            get
            {
                return Regex.Match(name, @"[A-Z]").Value;
            }
        }

        public uint RowIndex
        {
            get
            {
                return uint.Parse(Regex.Match(name, @"\d").Value);
            }
        }
    }
}
