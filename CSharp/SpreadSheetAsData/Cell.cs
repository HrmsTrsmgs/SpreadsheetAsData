using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Packaging = DocumentFormat.OpenXml.Packaging;
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

        internal Spreadsheet.Cell Xml { get;　private set; }

        public string Reference
        {
            get
            {
                return Xml.CellReference;
            }
        }
        public dynamic Value
        {
            get
            {
                switch (Xml.DataType)
                {
                    case "b":
                        return Xml.CellValue.Text != "0";
                    case "s":
                        return Book.Document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<Spreadsheet.SharedStringItem>().ElementAt(int.Parse(Xml.CellValue.Text)).Text.Text;
                    default:
                        return double.Parse(Xml.CellValue.Text);
                }   
            }
        }

        public Worksheet Sheet
        {
            get;
            private set;
        }
        public Workbook Book
        {
            get
            {
                return Sheet.Book;
            }
        }
    }
}
