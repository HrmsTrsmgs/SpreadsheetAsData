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

        private CellName(string name)
        {
            if(!Regex.IsMatch(name, @"[A-Z]\d"))
            {
                throw new FormatException();
            }
            this.name = name;
        }

        public uint ColumnIndex
        {
            get
            {
                return GetColumnIndex(ColumnName);
            }
        }

        private uint GetColumnIndex(IEnumerable<char> columnNameChars)
        {
            var count = columnNameChars.Count();
            switch(count)
            {
                case 1:
                    return (uint)(columnNameChars.Single() - 'A') + 1;
                default:
                    return 
                        GetColumnIndex(columnNameChars.Take(count - 1)) * 26
                          + GetColumnIndex(columnNameChars.Skip(count - 1));
            }
        }


        public static CellName Parse(string name)
        {
            return new CellName(name);
        }

        public string ColumnName
        {
            get
            {
                return Regex.Match(name, @"[A-Z]+").Value;
            }
        }

        public uint RowIndex
        {
            get
            {
                return uint.Parse(Regex.Match(name, @"\d").Value);
            }
        }

        public override string ToString()
        {
            return name;
        }
    }
}
