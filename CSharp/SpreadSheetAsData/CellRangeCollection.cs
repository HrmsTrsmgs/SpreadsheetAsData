using System.Collections.Generic;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ワークシート上のセル範囲を取得するコレクションです。
    /// </summary>
    public class CellRangeCollection
    {
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
    }
}
