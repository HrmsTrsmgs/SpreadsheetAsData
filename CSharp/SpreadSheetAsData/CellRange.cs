using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAsData
{
    /// <summary>
    /// ワークシート上のセル範囲を表します。
    /// </summary>
    public class CellRange
    {
        /// <summary>
        /// <see cref="ToString"/> で A1 形式の範囲を復元するための右下セル参照です。
        /// </summary>
        readonly string bottomRight;

        /// <summary>
        /// <see cref="ToString"/> で A1 形式の範囲を復元するための左上セル参照です。
        /// </summary>
        readonly string topLeft;

        /// <summary>
        /// 指定した左上セルと右下セルでセル範囲を作成します。
        /// </summary>
        /// <param name="topLeft">範囲の左上セル参照。</param>
        /// <param name="bottomRight">範囲の右下セル参照。</param>
        public CellRange(string topLeft, string bottomRight)
        {
            this.topLeft = topLeft;
            this.bottomRight = bottomRight;
        }

        /// <summary>
        /// 名前付き範囲として取得された場合の名前を取得します。
        /// </summary>
        public string? Name => null;

        /// <summary>
        /// 範囲の左上セルを取得します。
        /// </summary>
        public Cell TopLeftCell =>
            throw new NotImplementedException();

        /// <summary>
        /// 範囲の右下セルを取得します。
        /// </summary>
        public Cell BottomRightCell =>
            throw new NotImplementedException();

        /// <summary>
        /// A1 形式のセル範囲を返します。
        /// </summary>
        /// <returns>A1 形式のセル範囲。</returns>
        public override string ToString() => $"{topLeft}:{bottomRight}";
    }
}
