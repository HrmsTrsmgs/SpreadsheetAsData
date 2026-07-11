using System;
using System.Collections.Generic;

using System.Linq;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ワークシート上のセル範囲を取得するコレクションです。
    /// </summary>
    public class CellRangeCollection
    {
        /// <summary>
        /// ブックスコープの名前参照を解決するためのブックです。
        /// </summary>
        readonly Workbook? book;

        /// <summary>
        /// 同じ範囲指定に対して同じ <see cref="CellRange"/> インスタンスを返すためのキャッシュです。
        /// </summary>
        readonly Dictionary<(string TopLeft, string BottomRight), CellRange> cache = new();

        /// <summary>
        /// セル範囲が属するワークシートです。
        /// </summary>
        readonly Worksheet? sheet;

        /// <summary>
        /// 親を持たないセル範囲コレクションを作成します。
        /// </summary>
        public CellRangeCollection()
        { }

        /// <summary>
        /// 指定したブック上のセル範囲コレクションを作成します。
        /// </summary>
        /// <param name="book">範囲参照を解決するブック。</param>
        internal CellRangeCollection(Workbook book)
        {
            this.book = book;
        }

        /// <summary>
        /// 指定したワークシート上のセル範囲コレクションを作成します。
        /// </summary>
        /// <param name="sheet">範囲参照を解決するワークシート。</param>
        internal CellRangeCollection(Worksheet sheet)
        {
            this.sheet = sheet;
        }

        /// <summary>
        /// 左上セルと右下セルを指定してセル範囲を取得します。
        /// </summary>
        /// <param name="topLeft">範囲の左上セル参照。</param>
        /// <param name="bottomRight">範囲の右下セル参照。</param>
        /// <returns>指定したセル範囲。</returns>
        public CellRange this[string topLeft, string bottomRight]
        {
            get
            {
                var key = (topLeft, bottomRight);
                if (!cache.TryGetValue(key, out var range))
                {
                    range = new CellRange(sheet, topLeft, bottomRight);
                    cache[key] = range;
                }

                return range;
            }
        }

        /// <summary>
        /// A1形式または名前による範囲参照からセル範囲を取得します。
        /// </summary>
        /// <param name="reference">解決する範囲参照。</param>
        /// <returns>指定した範囲参照が表すセル範囲。</returns>
        public CellRange this[string reference]
        {
            get
            {
                var cellReferences = reference.Split(':');
                if (cellReferences.Length != 2)
                {
                    return GetWorkbookNamedRange(reference);
                }

                var topLeft = cellReferences[0];
                var bottomRight = cellReferences[1];
                return this[topLeft, bottomRight];
            }
        }

        /// <summary>
        /// ブックスコープの定義名をセル範囲として解決します。
        /// </summary>
        /// <param name="name">解決する定義名。</param>
        /// <returns>定義名が表すセル範囲。</returns>
        CellRange GetWorkbookNamedRange(string name)
        {
            var definedName = book?.WorkbookPart.Workbook.DefinedNames?.Elements<Spreadsheet.DefinedName>()
                .Where(_ => _.Name == name && _.LocalSheetId == null)
                .SingleOrDefault()
                ?? throw new NotImplementedException();

            var reference = definedName.Text.Split('!');
            if (reference.Length != )
            {
                throw new NotImplementedException();
            }

            var targetSheet = book.Sheets[reference[0]];
            var cellReferences = reference[1].Replace("$", "").Split(':');
            if (cellReferences.Length != 2)
            {
                throw new NotImplementedException();
            }

            var topLeft = cellReferences[0];
            var bottomRight = cellReferences[1];
            return new(targetSheet, topLeft, bottomRight);
        }
    }
}
