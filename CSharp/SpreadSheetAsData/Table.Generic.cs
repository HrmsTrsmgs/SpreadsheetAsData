namespace Marimo.SpreadSheetAsData;

/// <summary>
/// Excel テーブルの各データ行を指定した型へ対応付けて列挙する型付きテーブルを表します。
/// </summary>
/// <typeparam name="T">各データ行を対応付ける型。</typeparam>
public sealed class Table<T> : IEnumerable<T>
{
    /// <summary>
    /// 型付き列挙の元になる非型付き Excel テーブルです。
    /// </summary>
    readonly Table source;

    /// <summary>
    /// 指定した非型付き Excel テーブルから型付きテーブルを作成します。
    /// </summary>
    /// <param name="source">型付き列挙の元になる Excel テーブル。</param>
    internal Table(Table source)
    {
        this.source = source;
    }

    /// <summary>
    /// Excel テーブルの各データ行を <typeparamref name="T"/> へ対応付けて列挙します。
    /// </summary>
    /// <returns>型付き行の列挙子。</returns>
    public IEnumerator<T> GetEnumerator()
    {
        _ = source;
        throw new NotImplementedException();
    }

    /// <summary>
    /// Excel テーブルの各データ行を <typeparamref name="T"/> へ対応付けて列挙します。
    /// </summary>
    /// <returns>型付き行の列挙子。</returns>
    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
