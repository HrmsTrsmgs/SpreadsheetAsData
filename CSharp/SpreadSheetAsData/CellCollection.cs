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
                var cellName = CellName.Parse(cellReference);

                if (!cache.ContainsKey(cellName))
                {
                    var cellXml =
                        from cell in sheet.WorksheetPart.Worksheet.Descendants<Spreadsheet.Cell>()
                        where cell.CellReference == cellReference
                        select cell;
                    if (cellXml.Any())
                    {
                        cache[cellName] = new Cell(sheet, cellXml.Single());
                    }
                    else
                    {
                        cache[cellName] = new Cell(sheet, cellReference);
                    }
                }
                return cache[cellName];
            }
        }

        public Cell this[int columnIndex, int rowIndex]
        {
            get
            {
                return this[GetCellReference(columnIndex, rowIndex)];
            }
        }

        private string GetCellReference(int columnNumber, int rowNumber)
        {
            return ((char)('A' - 1 + columnNumber)).ToString() + rowNumber;
        }
    }
}
