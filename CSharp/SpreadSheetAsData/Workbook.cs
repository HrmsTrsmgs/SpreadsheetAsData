using System;
using System.Linq;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    public class Workbook : IDisposable
    {
        bool disposedValue;
        public static Workbook Open(string filePath) =>
            new Workbook(Packaging.SpreadsheetDocument.Open(filePath, true));

        Workbook(Packaging.SpreadsheetDocument document)
        {
            Document = document;
        }

        internal Packaging.SpreadsheetDocument Document { get; }

        internal Packaging.WorkbookPart WorkbookPart =>
            Document.WorkbookPart ?? throw new InvalidOperationException();

        WorksheetCollection? sheets { get; set; }

        public WorksheetCollection Sheets =>
            sheets ??= new WorksheetCollection(
                        from sheet in (WorkbookPart.Workbook.Sheets ?? throw new InvalidOperationException()).Elements<Spreadsheet.Sheet>()
                        select new Worksheet(this, sheet.Name?.Value ?? throw new InvalidOperationException()));
        
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
