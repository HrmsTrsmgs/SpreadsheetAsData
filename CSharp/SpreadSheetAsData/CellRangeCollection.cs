using System;
using System.Collections.Generic;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ワークシート上のセル範囲を取得するコレクションです。
    /// </summary>
    public class CellRangeCollection
    {
        /// <summary>
        /// 同じ範囲指定に対して同じ <see cref="CellRange"/> インスタンスを返すためのキャッシュです。
        /// </summary>
        readonly Dictionary<(string TopLeft, string BottomRight), CellRange> cache = new();

        /// <summary>
        /// 左上セルと右下セルを指定してセル範囲を取得します。
        /// </summary>
        /// <param name="topLeft">範囲の左上セル参照。</param>
        /// <param name="bottomRight">範囲の右下セル参照。</param>
        /// <returns>指定したセル範囲。</returns>
        public CellRange this[string topLeft, string bottomRight]
        {
            get
            {
                var key = (topLeft, bottomRight);
                if (!cache.TryGetValue(key, out var range))
                {
                    range = new CellRange(topLeft, bottomRight);
                    cache[key] = range;
                }

                return range;
            }
        }

        /// <summary>
        /// A1形式または名前による範囲参照からセル範囲を取得します。
        /// </summary>
        /// <param name="reference">解決する範囲参照。</param>
        /// <returns>指定した範囲参照が表すセル範囲。</returns>
        public CellRange this[string reference]
        {
            get
            {
                var range = reference.Split(':');
                return range.Length == 2
                    ? this[range[0], range[1]]
                    : throw new NotImplementedException();
            }
        }
    }
}
