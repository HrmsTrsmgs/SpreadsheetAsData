using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace Marimo.SpreadSheetAsData;
/// <summary>
/// Spreadsheet ファイルとして開いたブックを表します。
/// </summary>
public class Workbook : IDisposable
{
    /// <summary>
    /// Open XML ドキュメントと、元ファイルへの明示保存に必要な状態を保持します。
    /// </summary>
    readonly DocumentSession documentSession;

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
        new(DocumentSession.Open(filePath));

    /// <summary>
    /// 派生した型付きブックから、指定したファイルをブックとして開きます。
    /// </summary>
    /// <param name="filePath">開く Spreadsheet ファイルのパス。</param>
    protected Workbook(string filePath) : this(DocumentSession.Open(filePath))
    {
    }

    /// <summary>
    /// 既に開かれた Open XML ドキュメントのセッションを所有するブックを作成します。
    /// </summary>
    /// <param name="session">ブックとして扱う Open XML ドキュメントのセッション。</param>
    Workbook(DocumentSession session)
    {
        documentSession = session;
        Document = session.Document;
        Range = new(this);
        Cell = new(this);
        Tables = new(this);
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
    WorksheetCollection? sheets;

    /// <summary>
    /// ブックに含まれるワークシートの一覧を取得します。
    /// </summary>
    public WorksheetCollection Sheets =>
        sheets ??= new WorksheetCollection(
                    from sheet in (WorkbookPart.Workbook.Sheets ?? throw new InvalidOperationException()).Elements<Spreadsheet.Sheet>()
                    select new Worksheet(this, sheet.Name?.Value ?? throw new InvalidOperationException()));

    /// <summary>
    /// ブック内の定義名を列挙します。
    /// 名前付きセルや名前付き範囲を対象とし、Excelテーブル名は含みません。
    /// </summary>
    public IEnumerable<DefinedName> DefinedNames =>
        from definedName in WorkbookPart.Workbook.DefinedNames?.Elements<Spreadsheet.DefinedName>() ?? []
        let name = definedName.Name?.Value ?? throw new InvalidOperationException()
        let worksheet = DefinedNameWorksheet(definedName.LocalSheetId?.Value)
        select new DefinedName(
            name,
            worksheet,
            worksheet?.Range[name] ?? Range[name]);

    Worksheet? DefinedNameWorksheet(uint? localSheetId) =>
        localSheetId.HasValue
            ? Sheets[(int)localSheetId.Value]
            : null;

    /// <summary>
    /// ブックスコープの定義名をセル範囲として解決します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <returns>定義名が表すセル範囲。</returns>
    internal CellRange ResolveNamedRange(string name) =>
        ResolveNamedRange(name, localSheetId: null);

    /// <summary>
    /// ブックスコープの定義名をセル範囲として解決できるか確認します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <param name="range">定義名が見つかった場合のセル範囲。</param>
    /// <returns>定義名を解決できた場合は true。</returns>
    internal bool TryResolveNamedRange(string name, out CellRange range) =>
        TryResolveNamedRange(name, localSheetId: null, out range);

    /// <summary>
    /// 指定したワークシートスコープの定義名をセル範囲として解決します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <param name="localSheetId">定義名が属するワークシートの 0 始まりの位置。</param>
    /// <returns>定義名が表すセル範囲。</returns>
    internal CellRange ResolveNamedRange(string name, uint localSheetId) =>
        ResolveNamedRange(name, (uint?)localSheetId);

    /// <summary>
    /// 指定したワークシートスコープの定義名をセル範囲として解決できるか確認します。
    /// </summary>
    /// <param name="name">解決する定義名。</param>
    /// <param name="localSheetId">定義名が属するワークシートの 0 始まりの位置。</param>
    /// <param name="range">定義名が見つかった場合のセル範囲。</param>
    /// <returns>定義名を解決できた場合は true。</returns>
    internal bool TryResolveNamedRange(string name, uint localSheetId, out CellRange range) =>
        TryResolveNamedRange(name, (uint?)localSheetId, out range);

    CellRange ResolveNamedRange(string name, uint? localSheetId)
    {
        return TryResolveNamedRange(name, localSheetId, out var range)
            ? range
            : throw new NotImplementedException();
    }

    bool TryResolveNamedRange(string name, uint? localSheetId, out CellRange range)
    {
        var definedName = FindDefinedName(name, localSheetId);

        if (definedName == null)
        {
            range = default!;
            return false;
        }

        var rangeReference = CellRangeReference.Parse(definedName.Text);
        var targetSheet = Sheets[rangeReference.SheetName ?? throw new NotImplementedException()];
        range = new(
            targetSheet,
            rangeReference.TopLeft,
            rangeReference.BottomRight,
            name);
        return true;
    }

    Spreadsheet.DefinedName? FindDefinedName(string name, uint? localSheetId)
    {
        var definedNames = WorkbookPart.Workbook.DefinedNames?.Elements<Spreadsheet.DefinedName>();

        return definedNames == null
            ? null
            : (
                from definedName in definedNames
                where definedName.Name == name && HasLocalSheetId(definedName, localSheetId)
                select definedName
            ).SingleOrDefault();
    }

    static bool HasLocalSheetId(Spreadsheet.DefinedName definedName, uint? localSheetId)
    {
        return localSheetId.HasValue
            ? definedName.LocalSheetId?.Value == localSheetId.Value
            : definedName.LocalSheetId == null;
    }

    /// <summary>
    /// ブック上で有効な範囲参照を解決するコレクションを取得します。
    /// </summary>
    public CellRangeCollection Range { get; }

    /// <summary>
    /// ブック上で有効なセル参照を解決するコレクションを取得します。
    /// </summary>
    public CellCollection Cell { get; }

    /// <summary>
    /// ブック内の Excel テーブルを取得するコレクションを取得します。
    /// </summary>
    public TableCollection Tables { get; }

    /// <summary>
    /// 指定した名前の Excel テーブルを、各データ行を <typeparamref name="T"/> へ対応付ける型付きテーブルとして取得します。
    /// </summary>
    /// <typeparam name="T">各データ行を対応付ける型。</typeparam>
    /// <param name="name">取得する Excel テーブル名。</param>
    /// <returns>指定した Excel テーブルの型付き列挙。</returns>
    public Table<T> ReadTable<T>(string name) =>
        new(Tables[name]);

    /// <summary>
    /// 指定した名前の Excel テーブルへ、指定した型付き行をワークシート上の順序で書き込みます。
    /// </summary>
    /// <typeparam name="T">各データ行として書き込む型。</typeparam>
    /// <param name="name">書き込み先の Excel テーブル名。</param>
    /// <param name="rows">書き込む型付き行。</param>
    public void WriteTable<T>(string name, IEnumerable<T> rows) =>
        throw new NotImplementedException();

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
    public void Close() => documentSession.Dispose();

    /// <summary>
    /// ブックへの変更を、開いているファイルへ保存します。
    /// </summary>
    public void Save() =>
        documentSession.Save();

    /// <summary>
    /// ブックへの変更を、指定した別ファイルへ保存します。
    /// </summary>
    /// <param name="filePath">保存先のファイルパス。</param>
    public void SaveAs(string filePath) =>
        documentSession.SaveAs(filePath);

    /// <summary>
    /// ファイルをロックしたまま、編集対象の Open XML ドキュメントをメモリ上に保持します。
    /// </summary>
    sealed class DocumentSession : IDisposable
    {
        readonly string filePath;
        readonly MemoryStream stream;
        FileStream fileLock;
        bool disposedValue;

        DocumentSession(
            string filePath,
            FileStream fileLock,
            MemoryStream stream,
            Packaging.SpreadsheetDocument document)
        {
            this.filePath = filePath;
            this.fileLock = fileLock;
            this.stream = stream;
            Document = document;
        }

        internal Packaging.SpreadsheetDocument Document { get; }

        internal static DocumentSession Open(string filePath)
        {
            var fileLock = Lock(filePath);
            var stream = new MemoryStream();

            try
            {
                fileLock.CopyTo(stream);
                stream.Position = 0;

                return new(
                    filePath,
                    fileLock,
                    stream,
                    Packaging.SpreadsheetDocument.Open(
                        stream,
                        isEditable: true,
                        new Packaging.OpenSettings { AutoSave = false }));
            }
            catch
            {
                stream.Dispose();
                fileLock.Dispose();
                throw;
            }
        }

        static FileStream Lock(string filePath) =>
            File.Open(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite);

        internal void Save()
        {
            fileLock.Dispose();
            try
            {
                SaveAs(filePath);
            }
            finally
            {
                fileLock = Lock(filePath);
            }
        }

        internal void SaveAs(string filePath)
        {
            using var document = Document.Clone(filePath);
        }

        public void Dispose()
        {
            if (!disposedValue)
            {
                Document.Close();
                stream.Dispose();
                fileLock.Dispose();
                disposedValue = true;
            }
        }
    }

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
