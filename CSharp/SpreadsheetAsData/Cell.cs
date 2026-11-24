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
    /// 空白セルは <see cref="BlankValue"/>、真偽値セルは <see cref="bool"/>、共有文字列・インライン文字列・文字列セルは
    /// <see cref="string"/>、その他の数値セルは <see cref="double"/> として返します。
    /// 文字列は部分書式やふりがなを含まない本文として返します。
    /// 文字列を書き換える場合、セル内の部分書式は保持を保証しません。
    /// 値を設定すると、対象セルの既存の数式を削除します。
    /// また、Excelでブックを開く際に数式を再計算するよう要求します。ここでは計算結果を更新しません。
    /// </remarks>
    [System.Diagnostics.CodeAnalysis.AllowNull]
    public dynamic Value
    {
        // TypedValue の取得は空白も BlankValue で返すため、Union の中身は null になりません。
        get => TypedValue.Value!;
        set
        {
            object? input = value;
            TypedValue = NormalizeValue(input);
        }
    }

    /// <summary>
    /// 型を限定しない書き込み入力を、XMLを変更する前にセル値へ変換します。
    /// </summary>
    internal static TypedValue NormalizeValue(object? value)
    {
        return value switch
        {
            null => new BlankValue(),
            string text => text,
            bool boolean => boolean,
            _ => ToNumber(value)
        };

        static double ToNumber(object value)
        {
            // 明示的に渡された BlankValue も従来どおり数値へ変換します。
            dynamic input = value;
            return (double)input;
        }
    }

    /// <summary>
    /// セルの値を文字列、数値、真偽値、空白のいずれかとして取得または設定します。
    /// </summary>
    /// <remarks><see cref="Value"/> と同じセル値を扱い、読み書きの副作用も同じです。</remarks>
    public TypedValue TypedValue
    {
        get => (Xml.DataType?.Value, Xml.CellValue?.Text) switch
        {
            (null, null) => new BlankValue(),
            (var type, "0") when type == CellValues.Boolean => false,
            (var type, _) when type == CellValues.Boolean => true,
            (var type, _) when type == CellValues.SharedString => SharedStringValue,
            (var type, _) when type == CellValues.InlineString => StringValue(Xml.InlineString ?? throw new InvalidOperationException()),
            (var type, string text) when type == CellValues.String => text,
            (_, string text) => double.Parse(text, CultureInfo.InvariantCulture),
            _ => throw new InvalidOperationException()
        };


        set
        {
            (CellValues? DataType, string? Text) encoded = value switch
            {
                BlankValue => (null, null),
                string text => (CellValues.String, text),
                bool boolean => (CellValues.Boolean, boolean ? "1" : "0"),
                double number => (CellValues.Number, number.ToString(CultureInfo.InvariantCulture))
            };

            AttachToWorksheet();
            Xml.CellFormula = null;
            (Book.WorkbookXml.CalculationProperties ??= new()).FullCalculationOnLoad = true;
            Xml.DataType = encoded.DataType is CellValues type ? new(type) : null;
            Xml.CellValue = encoded.Text is null ? null : new(encoded.Text);
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

        var existingCell = row.Elements<Spreadsheet.Cell>().SingleOrDefault(it => it.CellReference?.Value == Reference);
        if (existingCell is not null)
        {
            Xml = existingCell;
            return;
        }

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

            return StringValue(sharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(text)));
        }
    }

    /// <summary>
    /// 直接の本文と本文runだけを連結し、ふりがなを除外します。
    /// </summary>
    static string StringValue(RstType text) =>
        string.Concat(
            from child in text.ChildElements
            where child is Text or Run
            select child.InnerText);

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
