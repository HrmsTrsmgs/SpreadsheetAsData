using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class Worksheet
    {
        public Worksheet()
        {
            Cells = new CellCollection(this);
        }

        public Workbook Book { get; internal set; }
        public string Name { get; internal set; }
        public CellCollection Cells { get; }
        public CellRangeCollection Range { get; } = new CellRangeCollection();

        internal Spreadsheet.Sheet SheetTag =>
            Book.Document.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>().Where(_ => _.Name == Name).Single();

        internal Packaging.WorksheetPart WorksheetPart =>
            Book.Document.WorkbookPart.GetPartById(SheetTag.Id) as Packaging.WorksheetPart;
    }
}
