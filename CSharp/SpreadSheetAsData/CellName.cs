using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// A1 形式のセル参照を表します。
    /// </summary>
    public partial struct CellName
    {
        /// <summary>
        /// Excel ワークシートで使用できる最大行番号です。
        /// </summary>
        public const uint MaxRowIndex = 1048576;

        /// <summary>
        /// Excel ワークシートで使用できる最大列番号です。
        /// </summary>
        public const uint MaxColumnIndex = 16384;
        const uint alphabetCount = 26;
        static Regex CellNamePattern => GeneratedCellNameRegex();

        /// <summary>
        /// 1 始まりの列番号を取得します。
        /// </summary>
        public uint ColumnIndex { get; private set; }

        /// <summary>
        /// 1 始まりの行番号を取得します。
        /// </summary>
        public uint RowIndex { get; private set; }

        CellName(string name)
        {
            var match = CellNamePattern.Match(name);
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

        /// <summary>
        /// 列番号と行番号からセル参照を作成します。
        /// </summary>
        /// <param name="columnIndex">1 始まりの列番号。</param>
        /// <param name="rowIndex">1 始まりの行番号。</param>
        /// <exception cref="FormatException">列番号または行番号が使用可能範囲を超えています。</exception>
        public CellName(uint columnIndex, uint rowIndex)
        {
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            if (MaxRowIndex < RowIndex || MaxColumnIndex < ColumnIndex)
            {
                throw new FormatException();
            }
        }

        /// <summary>
        /// A1 形式の文字列をセル参照に変換します。
        /// </summary>
        /// <param name="name">A1 形式のセル参照。</param>
        /// <returns>変換したセル参照。</returns>
        /// <exception cref="FormatException">文字列がA1形式でない、または使用可能範囲を超えています。</exception>
        public static CellName Parse(string name) => new CellName(name);

        /// <summary>
        /// 列名を取得します。
        /// </summary>
        public string ColumnName => GetColumnName(ColumnIndex);

        /// <summary>
        /// A1 形式のセル参照を返します。
        /// </summary>
        /// <returns>A1 形式のセル参照。</returns>
        public override string ToString() =>
            $"{GetColumnName(ColumnIndex)}{RowIndex}";

        static uint GetColumnIndex(IEnumerable<char> columnNameChars) =>
            columnNameChars.Count() switch
            {
                1 => (uint)(columnNameChars.Single() - 'A') + 1,
                _ => GetColumnIndex(columnNameChars.Take(columnNameChars.Count() - 1)) * alphabetCount
                          + GetColumnIndex(columnNameChars.Skip(columnNameChars.Count() - 1))
            };
        static string GetColumnName(uint columnIndex) =>
            (columnIndex <= alphabetCount) switch
            {
                true => ((char)('A' + columnIndex - 1)).ToString(),
                false => $"{GetColumnName((columnIndex - 1) / alphabetCount)}{GetColumnName((columnIndex - 1) % alphabetCount + 1)}"
            };

        [GeneratedRegex(@"^(?<column>[A-Z]+)(?<row>\d+)$")]
        private static partial Regex GeneratedCellNameRegex();
    }
}
