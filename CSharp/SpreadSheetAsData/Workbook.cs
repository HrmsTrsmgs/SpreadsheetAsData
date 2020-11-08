using System;
using System.Linq;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class Workbook : IDisposable
    {
        private bool disposedValue;
        public static Workbook Open(string filePath) =>
            new Workbook { Document = Packaging.SpreadsheetDocument.Open(filePath, true) };

        internal Packaging.SpreadsheetDocument Document { get; set; }

        WorksheetCollection sheets { get; set; }

        public WorksheetCollection Sheets =>
            sheets ??= new WorksheetCollection(
                        from sheet in Document.WorkbookPart.Workbook.Sheets.Elements<Spreadsheet.Sheet>()
                        select new Worksheet { Book = this, Name = sheet.Name.Value });
        
        public Worksheet this[int index] => Sheets[index];

        public Worksheet this[string sheetName] => Sheets[sheetName];

        public void Close() => Document.Close();

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    Close();
                }
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
