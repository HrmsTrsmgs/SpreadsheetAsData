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
            Range = new CellRangeCollection();
        }


        public Workbook Book { get; internal set; }
        public string Name { get; internal set; }
        public CellCollection Cells { get;private set; }
        public CellRangeCollection Range { get; private set; }

        internal Spreadsheet.Sheet SheetTag
        {
            get
            {
                return Book.Document.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>().Where(_ => _.Name == Name).Single();
            }
        }

        internal Packaging.WorksheetPart WorksheetPart
        {
            get
            {
                return Book.Document.WorkbookPart.GetPartById(SheetTag.Id) as Packaging.WorksheetPart;
            }
        }
    }
}
