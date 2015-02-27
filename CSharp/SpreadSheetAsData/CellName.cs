using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    public struct CellName
    {
        public const uint MaxRowIndex = 1048576;
        public const uint MaxColumnIndex = 16384;
        private const uint alphabetCount = 26;
        private static readonly Regex regex = new Regex(@"^(?<column>[A-Z]+)(?<row>\d+)$");

        public uint ColumnIndex { get; private set; }

        public uint RowIndex { get; private set; }

        private CellName(string name)
        {
            var match = regex.Match(name);
            if(!match.Success)
            {
                throw new FormatException();
            }
            ColumnIndex = GetColumnIndex(match.Groups["column"].Value);
            RowIndex = uint.Parse(match.Groups["row"].Value);

            if (MaxRowIndex < RowIndex || MaxColumnIndex <  ColumnIndex)
            {
                throw new FormatException();
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
                return GetColumnName(ColumnIndex);
            }
        }

        public override string ToString()
        {
            return GetColumnName(ColumnIndex) + RowIndex;
        }

        private static uint GetColumnIndex(IEnumerable<char> columnNameChars)
        {
            var count = columnNameChars.Count();
            switch(count)
            {
                case 1:
                    return (uint)(columnNameChars.Single() - 'A') + 1;
                default:
                    return 
                        GetColumnIndex(columnNameChars.Take(count - 1)) * alphabetCount
                          + GetColumnIndex(columnNameChars.Skip(count - 1));
            }
        }

        private static string GetColumnName(uint columnIndex)
        {
            if (columnIndex <= alphabetCount)
            {
                return ((char)('A' + columnIndex - 1)).ToString();
            }
            else
            {
                return GetColumnName(columnIndex / alphabetCount) + GetColumnName(columnIndex % alphabetCount);
            }
        }
    }
}
