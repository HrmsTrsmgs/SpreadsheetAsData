using Packaging = DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Validation = DocumentFormat.OpenXml.Validation;
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
    /// ブック内の定義名とExcelテーブルをデータオブジェクトへ対応付けます。
    /// </summary>
    readonly WorkbookDataMapper dataMapper;

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
    /// 指定したファイルをブックとして開きます。
    /// </summary>
    /// <param name="filePath">開く Spreadsheet ファイルのパス。</param>
    /// <param name="validate">Open XMLとして検証する場合はtrue。</param>
    /// <returns>開いたブック。</returns>
    public static Workbook Open(string filePath, bool validate) =>
        ValidateIfRequested(Open(filePath), validate);

    /// <summary>
    /// 指定したストリーム上の Spreadsheet ファイルをブックとして開きます。
    /// </summary>
    /// <param name="stream">開く Spreadsheet ファイルを格納したストリーム。</param>
    /// <returns>開いたブック。</returns>
    /// <remarks>
    /// 元Streamの内容は変更せず、CloseまたはDisposeでも元Stream自体は閉じません。
    /// Saveは使用できません。編集結果を出力する場合はSaveAsでファイルへ保存してください。
    /// </remarks>
    public static Workbook Open(Stream stream) =>
        new(DocumentSession.Open(stream));

    /// <summary>
    /// 派生した型付きブックから、指定したファイルをブックとして開きます。
    /// </summary>
    /// <param name="filePath">開く Spreadsheet ファイルのパス。</param>
    protected Workbook(string filePath) : this(DocumentSession.Open(filePath))
    {
    }

    /// <summary>
    /// 派生した型付きブックから、指定したストリーム上のファイルを開きます。
    /// </summary>
    /// <param name="stream">開く Spreadsheet ファイルを格納したストリーム。</param>
    protected Workbook(Stream stream) : this(DocumentSession.Open(stream))
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
        dataMapper = new(this);
    }

    static Workbook ValidateIfRequested(Workbook opened, bool validate)
    {
        if (!validate || !new Validation.OpenXmlValidator().Validate(opened.Document).Any())
        {
            return opened;
        }

        opened.Dispose();
        throw new InvalidDataException();
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

    CellRange ResolveNamedRange(string name, uint? localSheetId) =>
        TryResolveNamedRange(name, localSheetId, out var range)
            ? range
            : throw new KeyNotFoundException();

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
    public WorkbookCellRangeCollection Range { get; }

    /// <summary>
    /// ブック上で有効なセル参照を解決するコレクションを取得します。
    /// </summary>
    public WorkbookCellCollection Cell { get; }

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
    /// ブック内の定義名とExcelテーブルを、同じC#識別子となるプロパティへ対応付けて読み込みます。
    /// <see cref="SpreadSheetNameAttribute"/> による明示的な対応付けを優先します。
    /// </summary>
    /// <typeparam name="T">ブックのデータを読み込む型。</typeparam>
    /// <returns>ブックのデータを読み込んだオブジェクト。</returns>
    /// <remarks>
    /// 手書きの型でも、属性なしでExcel名をC#識別子へ自動変換して対応付けます。
    /// 例えば定義名cell_nameはCellNameへ対応します。シートローカル定義名も検索対象で、一意に対応する必要があります。
    /// </remarks>
    public T Read<T>() =>
        dataMapper.Read<T>();

    /// <summary>
    /// オブジェクトのプロパティを、同じC#識別子となるブック内の定義名またはExcelテーブルへ書き込みます。
    /// <see cref="SpreadSheetNameAttribute"/> による明示的な対応付けを優先します。
    /// </summary>
    /// <typeparam name="T">ブックへ書き込むデータの型。</typeparam>
    /// <param name="data">ブックへ書き込むデータ。</param>
    /// <remarks>
    /// 読み取りと同じ名前対応を使用します。例えば属性のないCellNameは、cell_nameという定義名へ書き込みます。
    /// シートローカル定義名を明示する場合は、属性のWorksheetNameも指定してください。
    /// </remarks>
    public void Replace<T>(T data) =>
        dataMapper.Replace(data);

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
    /// 変更を保存せず、ブックが使用しているリソースを解放します。
    /// </summary>
    /// <remarks>呼び出し側から渡されたStream自体は閉じません。</remarks>
    public void Close() => documentSession.Dispose();

    /// <summary>
    /// ブックへの変更を、開いているファイルへ保存します。
    /// </summary>
    /// <remarks>
    /// 正常終了時点で保存を完了します。CloseまたはDisposeを待つ必要はありません。
    /// Streamから開いた場合は元Streamを変更する前に拒否します。SaveAsで別ファイルへ保存してください。
    /// </remarks>
    /// <exception cref="NotSupportedException">Streamから開いたブックの場合。</exception>
    /// <exception cref="ObjectDisposedException">ファイルから開いたブックを既に閉じている場合。</exception>
    public void Save() =>
        documentSession.Save();

    /// <summary>
    /// ブックへの変更を、指定した別ファイルへ保存します。
    /// </summary>
    /// <param name="filePath">保存先のファイルパス。</param>
    /// <remarks>Streamから開いた場合も利用できます。元Streamの内容は変更しません。</remarks>
    public void SaveAs(string filePath) =>
        documentSession.SaveAs(filePath);

    /// <summary>
    /// ファイルまたはストリーム上の Open XML ドキュメントと保存先を管理します。
    /// </summary>
    sealed class DocumentSession : IDisposable
    {
        readonly string? filePath;

        /// <summary>
        /// SDKによるZIPの書き直しを元データへ漏らさない作業用Streamです。
        /// </summary>
        readonly CopyOnWriteStream stream;
        FileStream? fileLock;
        bool disposedValue;

        DocumentSession(
            string? filePath,
            FileStream? fileLock,
            CopyOnWriteStream stream,
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
            filePath = Path.GetFullPath(filePath);
            var fileLock = Lock(filePath);
            var stream = new CopyOnWriteStream(fileLock);

            try
            {
                // ファイル版は、Open時点の内容を元ファイルと切り離して保持します。
                stream.CreateWorkingCopy();

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

        internal static DocumentSession Open(Stream stream)
        {
            var documentStream = new CopyOnWriteStream(stream);
            try
            {
                return new(
                    filePath: null,
                    fileLock: null,
                    documentStream,
                    Packaging.SpreadsheetDocument.Open(
                        documentStream,
                        isEditable: true,
                        new Packaging.OpenSettings { AutoSave = false }));
            }
            catch
            {
                documentStream.Dispose();
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
            if (filePath is null)
            {
                throw new NotSupportedException();
            }

            ObjectDisposedException.ThrowIf(disposedValue, this);
            fileLock?.Dispose();
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
                fileLock?.Dispose();
                disposedValue = true;
            }
        }
    }

    /// <summary>
    /// 元のStreamを借りて読み込み、最初の変更から拡張可能なコピーへ切り替えます。
    /// コピーだけを所有し、元のStreamへの書き込みと破棄は行いません。
    /// </summary>
    sealed class CopyOnWriteStream : Stream
    {
        readonly Stream source;

        /// <summary>
        /// Openへ渡された時点の位置を、文書の先頭として扱います。
        /// </summary>
        readonly long sourceOffset;
        MemoryStream? workingCopy;

        internal CopyOnWriteStream(Stream source)
        {
            this.source = source;
            if (source.CanSeek)
            {
                sourceOffset = source.Position;
            }
            else
            {
                // SDKがランダムアクセスできるよう、シーク不可の場合は先にコピーします。
                CreateWorkingCopy();
            }
        }

        Stream Current => workingCopy ?? source;

        public override bool CanRead => Current.CanRead;
        public override bool CanSeek => Current.CanSeek;
        public override bool CanWrite => workingCopy?.CanWrite ?? true;
        public override long Length => workingCopy?.Length ?? source.Length - sourceOffset;

        public override long Position
        {
            get => workingCopy?.Position ?? source.Position - sourceOffset;
            set => Current.Position = workingCopy is null ? sourceOffset + value : value;
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            Current.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) =>
            workingCopy is not null
                ? workingCopy.Seek(offset, origin)
                : source.Seek(origin == SeekOrigin.Begin ? sourceOffset + offset : offset, origin) - sourceOffset;

        public override void Write(byte[] buffer, int offset, int count) =>
            CreateWorkingCopy().Write(buffer, offset, count);

        public override void SetLength(long value) =>
            CreateWorkingCopy().SetLength(value);

        public override void Flush() => workingCopy?.Flush();

        /// <summary>
        /// 読み書き位置を維持したままコピーへ切り替え、以降は同じコピーを使います。
        /// </summary>
        internal MemoryStream CreateWorkingCopy()
        {
            if (workingCopy is null)
            {
                var position = source.CanSeek ? Position : 0;
                if (source.CanSeek)
                {
                    source.Position = sourceOffset;
                }

                var copy = new MemoryStream();
                try
                {
                    source.CopyTo(copy);
                    copy.Position = position;
                    workingCopy = copy;
                }
                catch
                {
                    copy.Dispose();
                    throw;
                }
            }

            return workingCopy;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                workingCopy?.Dispose();
            }
            base.Dispose(disposing);
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
    /// 変更を保存せず、ブックが使用しているリソースを解放します。
    /// </summary>
    /// <remarks>呼び出し側から渡されたStream自体は閉じません。</remarks>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
