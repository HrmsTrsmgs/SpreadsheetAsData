using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Packaging =  DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class Workbook : IDisposable
    {
        public static Workbook Open(string filePath)
        {
            return new Workbook { Document = Packaging.SpreadsheetDocument.Open(filePath, true) };
        }

        internal Packaging.SpreadsheetDocument Document { get; set; }

        WorksheetCollection sheets { get; set; }

        public WorksheetCollection Sheets
        {
            get
            {
                if(sheets == null)
                {
                    sheets = new WorksheetCollection(
                        from sheet in Document.WorkbookPart.Workbook.Sheets.Elements<Spreadsheet.Sheet>()
                        select new Worksheet { Book = this, Name = sheet.Name.Value });
                }
                return sheets;
            }
        }

        public Worksheet this[int index] => Sheets[index];

        public Worksheet this[string sheetName] => Sheets[sheetName];

        public void Close()
        {
            Document.Close();
        }

        public void Dispose()
        {
            Close();
        }
    }
}
