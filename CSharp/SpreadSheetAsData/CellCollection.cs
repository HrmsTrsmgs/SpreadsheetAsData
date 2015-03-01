using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class CellCollection
    {
        Worksheet sheet;
        public CellCollection(Worksheet sheet)
        {
            this.sheet = sheet;
        }

        private Dictionary<CellName, Cell> cache = new Dictionary<CellName, Cell>();

        public Cell this[string cellReference]
        {
            get
            {
                return GetItem(CellName.Parse(cellReference));
            }
        }

        public Cell this[uint columnIndex, uint rowIndex]
        {
            get
            {
                return GetItem(new CellName(columnIndex, rowIndex));
            }
        }

        private Cell GetItem(CellName cellName)
        {
            if (!cache.ContainsKey(cellName))
            {
                var cellXml =
                    from cell in sheet.WorksheetPart.Worksheet.Descendants<Spreadsheet.Cell>()
                    where cell.CellReference == cellName.ToString()
                    select cell;
                if (cellXml.Any())
                {
                    cache[cellName] = new Cell(sheet, cellXml.Single());
                }
                else
                {
                    cache[cellName] = new Cell(sheet, cellName.ToString());
                }
            }
            return cache[cellName];
        }
    }
}
