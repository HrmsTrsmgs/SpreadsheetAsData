namespace Marimo.SpreadSheetAsData;

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
