using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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
            if (!match.Success)
            {
                throw new FormatException();
            }
            ColumnIndex = GetColumnIndex(match.Groups["column"].Value);
            RowIndex = uint.Parse(match.Groups["row"].Value);

            if (MaxRowIndex < RowIndex || MaxColumnIndex < ColumnIndex)
            {
                throw new FormatException();
            }
        }

        public CellName(uint columnIndex, uint rowIndex)
        {
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            if (MaxRowIndex < RowIndex || MaxColumnIndex < ColumnIndex)
            {
                throw new FormatException();
            }
        }

        public static CellName Parse(string name) => new CellName(name);

        public string ColumnName => GetColumnName(ColumnIndex);

        public override string ToString() =>
            GetColumnName(ColumnIndex) + RowIndex;

        private static uint GetColumnIndex(IEnumerable<char> columnNameChars) =>
            columnNameChars.Count() switch
            {
                1 => (uint)(columnNameChars.Single() - 'A') + 1,
                _ => GetColumnIndex(columnNameChars.Take(columnNameChars.Count() - 1)) * alphabetCount
                          + GetColumnIndex(columnNameChars.Skip(columnNameChars.Count() - 1))
            };
        private static string GetColumnName(uint columnIndex) =>
            (columnIndex <= alphabetCount) switch
            {
                true => ((char)('A' + columnIndex - 1)).ToString(),
                false => GetColumnName((columnIndex - 1) / alphabetCount) + GetColumnName((columnIndex - 1) % alphabetCount + 1)
            };
    }
}
