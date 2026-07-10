using System.Collections.Generic;
using System.Linq;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ワークシート上のセルを取得するコレクションです。
    /// </summary>
    public class CellCollection
    {
        /// <summary>
        /// Open XML のセル探索と空白セル作成に使用するワークシートです。
        /// </summary>
        readonly Worksheet sheet;

        /// <summary>
        /// 同じセル参照に対して同じ <see cref="Cell"/> インスタンスを返すためのキャッシュです。
        /// </summary>
        readonly Dictionary<CellName, Cell> cache = new();

        /// <summary>
        /// 指定したワークシートのセルコレクションを作成します。
        /// </summary>
        /// <param name="sheet">対象のワークシート。</param>
        public CellCollection(Worksheet sheet)
        {
            this.sheet = sheet;
        }

        /// <summary>
        /// A1 形式のセル参照でセルを取得します。
        /// </summary>
        /// <param name="cellReference">A1 形式のセル参照。</param>
        /// <returns>指定したセル。</returns>
        public Cell this[string cellReference] =>
            GetItem(CellName.Parse(cellReference));

        /// <summary>
        /// 列番号と行番号でセルを取得します。
        /// </summary>
        /// <param name="columnIndex">1 始まりの列番号。</param>
        /// <param name="rowIndex">1 始まりの行番号。</param>
        /// <returns>指定したセル。</returns>
        public Cell this[uint columnIndex, uint rowIndex] =>
            GetItem(new CellName(columnIndex, rowIndex));

        /// <summary>
        /// キャッシュ、既存の Open XML セル、空白セルの順でセルを解決します。
        /// </summary>
        /// <param name="cellName">取得するセル参照。</param>
        /// <returns>指定したセル。</returns>
        Cell GetItem(CellName cellName)
        {
            if (cache.TryGetValue(cellName, out var cachedCell))
            {
                return cachedCell;
            }

            var cellReference = cellName.ToString();
            var cellXml =
                from xml in sheet.WorksheetPart.Worksheet.Descendants<Spreadsheet.Cell>()
                where xml.CellReference == cellReference
                select xml;

            var cell = cellXml.Any()
                ? new Cell(sheet, cellXml.Single())
                : new Cell(sheet, cellReference);

            cache[cellName] = cell;
            return cell;
        }
    }
}
