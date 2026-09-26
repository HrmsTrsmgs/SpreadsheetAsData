using DocumentFormat.OpenXml.Packaging;
using Packaging = DocumentFormat.OpenXml.Packaging;

namespace Marimo.SpreadsheetAsData;

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
                OpenDocument(stream));
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
                OpenDocument(documentStream));
        }
        catch
        {
            documentStream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 空の入力を拒否し、明示的な保存だけを行うSDKドキュメントを開きます。
    /// </summary>
    static Packaging.SpreadsheetDocument OpenDocument(CopyOnWriteStream stream)
    {
        if (stream.Length == 0)
        {
            throw new InvalidDataException();
        }

        return Packaging.SpreadsheetDocument.Open(
            stream,
            isEditable: true,
            new Packaging.OpenSettings { AutoSave = false });
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

    /// <summary>
    /// 編集中の文書を出力先へ複製し、複製側を閉じて出力を完了します。
    /// </summary>
    internal void SaveAs(Stream destination)
    {
        if (ReferenceEquals(destination, stream.Source) || !destination.CanWrite)
        {
            throw new NotSupportedException();
        }

        using var document = Document.Clone(destination);
    }

    public void Dispose()
    {
        if (!disposedValue)
        {
            Document.Dispose();
            stream.Dispose();
            fileLock?.Dispose();
            disposedValue = true;
        }
    }
}
