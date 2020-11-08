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

        internal Spreadsheet.Row RowXml => Xml.Parent as Row;

        public string Reference => Xml.CellReference;

        public dynamic Value =>
            (Xml.DataType?.Value, Xml.CellValue?.Text) switch
            {
                (null, null) => new BlankValue(),
                (CellValues.Boolean, "0") => false,
                (CellValues.Boolean, _) => true,
                (CellValues.SharedString, _) =>
                    Book.Document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(Xml.CellValue.Text)).Text.Text,
                _ => double.Parse(Xml.CellValue.Text)
            };
        
        public Worksheet Sheet { get; private set; }

        public Workbook Book => Sheet.Book;

        public uint RowIndex => RowXml.RowIndex;

        public uint ColumnIndex => CellName.Parse(Reference).ColumnIndex;
    }
}
