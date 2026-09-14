using Packaging = DocumentFormat.OpenXml.Packaging;

namespace Marimo.SpreadSheetAsData;

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
            FileShare.Read);

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
        ObjectDisposedException.ThrowIf(disposedValue, this);
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
