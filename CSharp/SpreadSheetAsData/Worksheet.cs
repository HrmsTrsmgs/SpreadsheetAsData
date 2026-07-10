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
        private readonly Workbook? book;
        private readonly string? name;

        public Worksheet()
        {
            Cells = new CellCollection(this);
        }

        internal Worksheet(Workbook book, string name) : this()
        {
            this.book = book;
            this.name = name;
        }

        public Workbook Book =>
            book ?? throw new InvalidOperationException();

        public string Name =>
            name ?? throw new InvalidOperationException();

        public CellCollection Cells { get; }
        public CellRangeCollection Range { get; } = new CellRangeCollection();

        internal Spreadsheet.Sheet SheetTag =>
            Book.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>().Where(_ => _.Name == Name).Single();

        internal Packaging.WorksheetPart WorksheetPart =>
            Book.WorkbookPart.GetPartById(SheetTag.Id?.Value ?? throw new InvalidOperationException()) as Packaging.WorksheetPart
                ?? throw new InvalidOperationException();
    }
}
