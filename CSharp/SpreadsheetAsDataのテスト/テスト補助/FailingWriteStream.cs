namespace Marimo.SpreadsheetAsData.Test.テスト補助;

/// <summary>
/// 指定バイト数までは書き込み、それを超える出力をIOExceptionで中断します。
/// </summary>
sealed class FailingWriteStream(int bytesBeforeFailure) : MemoryStream
{
    /// <summary>
    /// 出力開始前ではなく、実際の書き込み途中で失敗したことを確認するための累計です。
    /// </summary>
    internal int WrittenByteCount { get; private set; }

    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count)
    {
        var writableCount = Math.Min(count, bytesBeforeFailure - WrittenByteCount);
        base.Write(buffer, offset, writableCount);
        WrittenByteCount += writableCount;
        if (writableCount < count)
        {
            throw new IOException("テスト用の書き込み途中の失敗です。");
        }
    }
}
