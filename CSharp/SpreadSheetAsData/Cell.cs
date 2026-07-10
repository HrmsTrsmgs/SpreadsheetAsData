using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class Cell
    {
        public Cell(Worksheet sheet, Spreadsheet.Cell xml)
        {
            Sheet = sheet;
            Xml = xml;
        }

        public Cell(Worksheet sheet, string cellReference) :
            this(
                sheet,
                new Spreadsheet.Cell(new Value { })
                {
                    CellReference = new StringValue(cellReference)
                })
        { }

        internal Spreadsheet.Cell Xml { get;　private set; }

        internal Spreadsheet.Row RowXml =>
            Xml.Parent as Row ?? throw new InvalidOperationException();

        public string Reference =>
            Xml.CellReference?.Value ?? throw new InvalidOperationException();

        public dynamic Value =>
            (Xml.DataType?.Value, Xml.CellValue?.Text) switch
            {
                (null, null) => new BlankValue(),
                (CellValues.Boolean, "0") => false,
                (CellValues.Boolean, _) => true,
                (CellValues.SharedString, _) => SharedStringValue,
                (_, string text) => double.Parse(text),
                _ => throw new InvalidOperationException()
            };

        string SharedStringValue
        {
            get
            {
                var text = Xml.CellValue?.Text ?? throw new InvalidOperationException();
                var sharedStringTable =
                    Book.WorkbookPart.SharedStringTablePart?.SharedStringTable ?? throw new InvalidOperationException();

                return sharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(text)).Text?.Text
                    ?? throw new InvalidOperationException();
            }
        }
        
        public Worksheet Sheet { get; private set; }

        public Workbook Book => Sheet.Book;

        public uint RowIndex => RowXml.RowIndex;

        public uint ColumnIndex => CellName.Parse(Reference).ColumnIndex;
    }
}
