using System.Collections.Generic;
using System.Linq;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class CellCollection
    {
        readonly Worksheet sheet;
        readonly Dictionary<CellName, Cell> cache = new();

        public CellCollection(Worksheet sheet)
        {
            this.sheet = sheet;
        }

        public Cell this[string cellReference] =>
            GetItem(CellName.Parse(cellReference));

        public Cell this[uint columnIndex, uint rowIndex] =>
            GetItem(new CellName(columnIndex, rowIndex));

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
