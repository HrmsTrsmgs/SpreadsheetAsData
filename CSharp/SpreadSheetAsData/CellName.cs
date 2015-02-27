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

        public const uint MaxRowIndex = 1048576;
        public const uint MaxColumnIndex = 16384;
        private CellName(string name)
        {
            this.name = name;
            if (!Regex.IsMatch(name, @"^[A-Z]+\d+$") || MaxRowIndex < RowIndex || MaxColumnIndex <  ColumnIndex)
            {
                throw new FormatException();
            }

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
                return uint.Parse(Regex.Match(name, @"\d+").Value);
            }
        }

        public override string ToString()
        {
            return name;
        }

        public override int GetHashCode()
        {
            return name.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var cellName = obj as CellName;

            if(cellName == null)
            {
                return false;
            }

            return cellName.name == name;
        }
    }
}
