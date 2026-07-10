using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ワークシート上のセルを表します。
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Open XML のセル要素からセルを作成します。
        /// </summary>
        /// <param name="sheet">セルが属するワークシート。</param>
        /// <param name="xml">セルを表す Open XML 要素。</param>
        public Cell(Worksheet sheet, Spreadsheet.Cell xml)
        {
            Sheet = sheet;
            Xml = xml;
        }

        /// <summary>
        /// 指定したセル参照の空白セルを作成します。
        /// </summary>
        /// <param name="sheet">セルが属するワークシート。</param>
        /// <param name="cellReference">A1 形式のセル参照。</param>
        public Cell(Worksheet sheet, string cellReference) :
            this(
                sheet,
                new Spreadsheet.Cell(new Value { })
                {
                    CellReference = new StringValue(cellReference)
                })
        { }

        /// <summary>
        /// このセルに対応する Open XML のセル要素を取得します。
        /// </summary>
        internal Spreadsheet.Cell Xml { get; private set; }

        /// <summary>
        /// このセルを含む Open XML の行要素を取得します。
        /// </summary>
        internal Spreadsheet.Row RowXml =>
            Xml.Parent as Row ?? throw new InvalidOperationException();

        /// <summary>
        /// A1 形式のセル参照を取得します。
        /// </summary>
        public string Reference =>
            Xml.CellReference?.Value ?? throw new InvalidOperationException();

        /// <summary>
        /// セルの値を取得します。
        /// </summary>
        /// <remarks>
        /// 空白セルは <see cref="BlankValue"/>、真偽値セルは <see cref="bool"/>、共有文字列セルは <see cref="string"/>、
        /// その他の数値セルは <see cref="double"/> として返します。
        /// </remarks>
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

        /// <summary>
        /// Open XML の共有文字列インデックスを実際の文字列へ解決します。
        /// </summary>
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

        /// <summary>
        /// セルが属するワークシートを取得します。
        /// </summary>
        public Worksheet Sheet { get; private set; }

        /// <summary>
        /// セルが属するブックを取得します。
        /// </summary>
        public Workbook Book => Sheet.Book;

        /// <summary>
        /// セルの行番号を取得します。
        /// </summary>
        public uint RowIndex => RowXml.RowIndex;

        /// <summary>
        /// セルの列番号を取得します。
        /// </summary>
        public uint ColumnIndex => CellName.Parse(Reference).ColumnIndex;
    }
}
