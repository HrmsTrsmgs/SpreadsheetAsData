using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadsheetAsData;
/// <summary>
/// ワークシート上のセルを表します。
/// </summary>
public class Cell
{
    /// <summary>
    /// Open XML から見つけた既存セルをラップします。
    /// </summary>
    /// <param name="sheet">セルが属するワークシート。</param>
    /// <param name="xml">セルを表す Open XML 要素。</param>
    internal Cell(Worksheet sheet, Spreadsheet.Cell xml)
    {
        Sheet = sheet;
        Xml = xml;
    }

    /// <summary>
    /// 指定したセル参照の空白セルを作成します。
    /// </summary>
    /// <param name="sheet">セルが属するワークシート。</param>
    /// <param name="cellReference">A1 形式のセル参照。</param>
    internal Cell(Worksheet sheet, string cellReference) :
        this(
            sheet,
            new Spreadsheet.Cell
            {
                CellReference = new StringValue(cellReference)
            })
    { }

    /// <summary>
    /// このセルに対応する Open XML のセル要素を取得します。
    /// </summary>
    internal Spreadsheet.Cell Xml { get; private set; }

    /// <summary>
    /// A1 形式のセル参照を取得します。
    /// </summary>
    public string Reference =>
        Xml.CellReference?.Value ?? throw new InvalidOperationException();

    /// <summary>
    /// セルの値を取得または設定します。
    /// </summary>
    /// <remarks>
    /// 空白セルは <see cref="BlankValue"/>、真偽値セルは <see cref="bool"/>、共有文字列セルと文字列セルは
    /// <see cref="string"/>、その他の数値セルは <see cref="double"/> として返します。
    /// 値を設定すると、対象セルの既存の数式を削除します。
    /// また、Excelでブックを開く際に数式を再計算するよう要求します。ここでは計算結果を更新しません。
    /// </remarks>
    [System.Diagnostics.CodeAnalysis.AllowNull]
    public dynamic Value
    {
        get => (Xml.DataType?.Value, Xml.CellValue?.Text) switch
        {
            (null, null) => new BlankValue(),
            (var type, "0") when type == CellValues.Boolean => false,
            (var type, _) when type == CellValues.Boolean => true,
            (var type, _) when type == CellValues.SharedString => SharedStringValue,
            (var type, string text) when type == CellValues.String => text,
            (_, string text) => double.Parse(text, CultureInfo.InvariantCulture),
            _ => throw new InvalidOperationException()
        };

        set
        {
            AttachToWorksheet();
            Xml.CellFormula = null;
            (Book.WorkbookXml.CalculationProperties ??= new()).FullCalculationOnLoad = true;

            if (value is null)
            {
                Xml.DataType = null;
                Xml.CellValue = null;
                return;
            }

            if (value is string text)
            {
                Xml.DataType = CellValues.String;
                Xml.CellValue = new(text);
                return;
            }

            if (value is bool boolean)
            {
                Xml.DataType = CellValues.Boolean;
                Xml.CellValue = new(boolean ? "1" : "0");
                return;
            }

            Xml.DataType = CellValues.Number;
            Xml.CellValue = new(((double)value).ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>
    /// 未格納のセルを初回書き込み時にXMLへ接続し、行と列の昇順を維持します。
    /// </summary>
    void AttachToWorksheet()
    {
        if (Xml.Parent is not null)
        {
            return;
        }

        var cellName = CellName.Parse(Reference);
        var sheetData = Sheet.WorksheetXml.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException();
        var row = sheetData.Elements<Row>().SingleOrDefault(it => RowIndexOf(it) == cellName.RowIndex)
            ?? sheetData.InsertBefore(
                new Row { RowIndex = cellName.RowIndex },
                sheetData.Elements<Row>().FirstOrDefault(it => RowIndexOf(it) > cellName.RowIndex));

        row.InsertBefore(
            Xml,
            row.Elements<Spreadsheet.Cell>().FirstOrDefault(
                it => CellName.Parse(it.CellReference?.Value ?? throw new NotSupportedException()).ColumnIndex > cellName.ColumnIndex));
    }

    static uint RowIndexOf(Row row)
    {
        if (row.RowIndex?.Value is uint index)
        {
            return index;
        }

        var indexes = (
            from cell in row.Elements<Spreadsheet.Cell>()
            let reference = cell.CellReference?.Value ?? throw new NotSupportedException()
            select CellName.Parse(reference).RowIndex
        ).Distinct().Take(2).ToArray();

        return indexes.Length == 1 ? indexes[0] : throw new NotSupportedException();
    }

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
    public uint RowIndex => CellName.Parse(Reference).RowIndex;

    /// <summary>
    /// セルの列番号を取得します。
    /// </summary>
    public uint ColumnIndex => CellName.Parse(Reference).ColumnIndex;

    /// <summary>
    /// A1 形式のセル参照を返します。
    /// </summary>
    /// <returns>A1 形式のセル参照。</returns>
    public override string ToString() => Reference;
}
