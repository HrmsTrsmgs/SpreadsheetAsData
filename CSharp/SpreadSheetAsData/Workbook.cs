using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// Spreadsheet ファイルとして開いたブックを表します。
/// </summary>
public class Workbook : IDisposable
{
    /// <summary>
    /// <see cref="Dispose(bool)"/> の多重実行を防ぐための状態です。
    /// </summary>
    bool disposedValue;

    /// <summary>
    /// 指定したファイルをブックとして開きます。
    /// </summary>
    /// <param name="filePath">開く Spreadsheet ファイルのパス。</param>
    /// <returns>開いたブック。</returns>
    public static Workbook Open(string filePath) =>
        new Workbook(Packaging.SpreadsheetDocument.Open(filePath, true));

    /// <summary>
    /// 既に開かれた Open XML ドキュメントを所有するブックを作成します。
    /// </summary>
    /// <param name="document">ブックとして扱う Open XML ドキュメント。</param>
    Workbook(Packaging.SpreadsheetDocument document)
    {
        Document = document;
        Range = new(this);
    }

    /// <summary>
    /// このブックが保持する Open XML ドキュメントを取得します。
    /// </summary>
    internal Packaging.SpreadsheetDocument Document { get; }

    /// <summary>
    /// このブックの Open XML ブックパートを取得します。
    /// </summary>
    internal Packaging.WorkbookPart WorkbookPart =>
        Document.WorkbookPart ?? throw new InvalidOperationException();

    /// <summary>
    /// Open XML のシート一覧から遅延作成したワークシートコレクションです。
    /// </summary>
    WorksheetCollection? sheets { get; set; }

    /// <summary>
    /// ブックに含まれるワークシートの一覧を取得します。
    /// </summary>
    public WorksheetCollection Sheets =>
        sheets ??= new WorksheetCollection(
                    from sheet in (WorkbookPart.Workbook.Sheets ?? throw new InvalidOperationException()).Elements<Spreadsheet.Sheet>()
                    select new Worksheet(this, sheet.Name?.Value ?? throw new InvalidOperationException()));

    /// <summary>
    /// ブックスコープの定義名をセル範囲として解決します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <returns>定義名が表すセル範囲。</returns>
    internal CellRange ResolveNamedRange(string name)
    {
        var definedName = WorkbookPart.Workbook.DefinedNames?.Elements<Spreadsheet.DefinedName>()
            .Where(_ => _.Name == name && _.LocalSheetId == null)
            .SingleOrDefault()
            ?? throw new NotImplementedException();

        var rangeReference = CellRangeReference.Parse(definedName.Text);
        var targetSheet = Sheets[rangeReference.SheetName ?? throw new NotImplementedException()];
        return new(targetSheet, rangeReference.TopLeft, rangeReference.BottomRight);
    }

    /// <summary>
    /// ブック上で有効な範囲参照を解決するコレクションを取得します。
    /// </summary>
    public CellRangeCollection Range { get; }

    /// <summary>
    /// 指定した位置のワークシートを取得します。
    /// </summary>
    /// <param name="index">取得するワークシートの 0 始まりの位置。</param>
    /// <returns>指定した位置のワークシート。</returns>
    public Worksheet this[int index] => Sheets[index];

    /// <summary>
    /// 指定した名前のワークシートを取得します。
    /// </summary>
    /// <param name="sheetName">取得するワークシート名。</param>
    /// <returns>指定した名前のワークシート。</returns>
    public Worksheet this[string sheetName] => Sheets[sheetName];

    /// <summary>
    /// ブックが保持しているファイルを閉じます。
    /// </summary>
    public void Close() => Document.Close();

    /// <summary>
    /// ブックが使用しているリソースを解放します。
    /// </summary>
    /// <param name="disposing">マネージドリソースを解放する場合は true。</param>
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

    /// <summary>
    /// ブックが使用しているリソースを解放します。
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
