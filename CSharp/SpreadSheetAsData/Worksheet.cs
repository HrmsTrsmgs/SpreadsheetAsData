using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ブック内のワークシートを表します。
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// ブックから作成されたワークシートだけが保持する親ブックです。
        /// </summary>
        readonly Workbook? book;

        /// <summary>
        /// ブックから作成されたワークシートだけが保持するシート名です。
        /// </summary>
        readonly string? name;

        /// <summary>
        /// 空のワークシートを作成します。
        /// </summary>
        public Worksheet()
        {
            Cells = new CellCollection(this);
            Range = new CellRangeCollection(this);
        }

        /// <summary>
        /// 指定したブック内のワークシートを作成します。
        /// </summary>
        /// <param name="book">ワークシートが属するブック。</param>
        /// <param name="name">ワークシート名。</param>
        internal Worksheet(Workbook book, string name) : this()
        {
            this.book = book;
            this.name = name;
        }

        /// <summary>
        /// このワークシートが属するブックを取得します。
        /// </summary>
        public Workbook Book =>
            book ?? throw new InvalidOperationException();

        /// <summary>
        /// ワークシート名を取得します。
        /// </summary>
        public string Name =>
            name ?? throw new InvalidOperationException();

        /// <summary>
        /// ワークシート上のセルを取得するコレクションを取得します。
        /// </summary>
        public CellCollection Cells { get; }

        /// <summary>
        /// ワークシート上のセル範囲を取得するコレクションを取得します。
        /// </summary>
        public CellRangeCollection Range { get; }

        /// <summary>
        /// このワークシートに対応する Open XML のシート要素を取得します。
        /// </summary>
        internal Spreadsheet.Sheet SheetTag =>
            Book.WorkbookPart.Workbook.Descendants<Spreadsheet.Sheet>().Where(_ => _.Name == Name).Single();

        /// <summary>
        /// このワークシートに対応する Open XML のワークシートパートを取得します。
        /// </summary>
        internal Packaging.WorksheetPart WorksheetPart =>
            Book.WorkbookPart.GetPartById(SheetTag.Id?.Value ?? throw new InvalidOperationException()) as Packaging.WorksheetPart
                ?? throw new InvalidOperationException();
    }
}
