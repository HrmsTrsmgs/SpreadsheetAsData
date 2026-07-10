using Marimo.SpreadSheetAsData;
using Xunit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Marimo.SpreadSheetAdData.Test
{
    public class CellNameのテスト
    {
        [Fact]
        public void ParseメソッドでCellNameが生成できます()
        {
            Assert.Null(Record.Exception(() => CellName.Parse("A1")));
        }

        [Fact]
        public void 列番号と行番号を指定してCellNameがせいせいできます()
        {
            Assert.Null(Record.Exception(() => new CellName(1,1)));
        }

        [Fact]
        public void 同じセル位置は同一として扱われます()
        {
            Assert.Equal(CellName.Parse("A1"), CellName.Parse("A1"));
        }

        [Fact]
        public void 行が違うと同一ではないとして扱われます()
        {
            Assert.NotEqual(CellName.Parse("A2"), CellName.Parse("A1"));
        }

        [Fact]
        public void 列が違うと同一ではないとして扱われます()
        {
            Assert.NotEqual(CellName.Parse("B1"), CellName.Parse("A1"));
        }

        [Fact]
        public void 指定したセルの位置を一意としてハッシュのキーとして使えます()
        {
            var set = new HashSet<CellName>();

            set.Add(CellName.Parse("A1"));
            Assert.Single(set);
            set.Add(CellName.Parse("A1"));
            Assert.Single(set);
            set.Add(CellName.Parse("B1"));
            Assert.Equal(2, set.Count);
            set.Add(CellName.Parse("A2"));
            Assert.Equal(3, set.Count);
        }

        [Fact]
        public void Parseメソッドに無効なセル名を指定するとFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("A"));
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("1"));
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("A1A1"));
        }

        [Fact]
        public void Parseメソッドに大きすぎる列名を指定した場合はFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("XFE1"));
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("AAAA1"));
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("ZZZZ1"));
        }

        [Fact]
        public void コンストラクタに大きすぎる列番号を指定した場合はFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => new CellName(16385, 1));
        }

        [Fact]
        public void Parseメソッドに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => CellName.Parse("A1048577"));
        }

        [Fact]
        public void コンストラクタに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
        {
            Assert.ThrowsAny<FormatException>(() => new CellName(1, 1048577));
        }

        [Fact]
        public void ColumnNameプロパティで列名を取得できます()
        {
            Assert.Equal("A", CellName.Parse("A1").ColumnName);
            Assert.Equal("B", CellName.Parse("B1").ColumnName);
        }

        [Fact]
        public void 行列番号で生成した場合にColumnNameプロパティで列名を取得できます()
        {
            Assert.Equal("A", new CellName(1, 1).ColumnName);
            Assert.Equal("B", new CellName(2, 1).ColumnName);
        }

        [Fact]
        public void ColumnIndexプロパティで一文字の列番号を取得できます()
        {
            Assert.Equal((uint)1, CellName.Parse("A1").ColumnIndex);
            Assert.Equal((uint)26, CellName.Parse("Z1").ColumnIndex);
        }

        [Fact]
        public void ColumnIndexプロパティで二文字のの列番号を取得できます()
        {
            Assert.Equal((uint)27, CellName.Parse("AA1").ColumnIndex);
            Assert.Equal((uint)702, CellName.Parse("ZZ1").ColumnIndex);
        }

        [Fact]
        public void ColumnIndexプロパティで三文字のの列番号を取得できます()
        {
            Assert.Equal((uint)703, CellName.Parse("AAA1").ColumnIndex);
            Assert.Equal((uint)16384, CellName.Parse("XFD1").ColumnIndex);
        }

        [Fact]
        public void コンストラクタで生成した場合にColumnIndexプロパティで大きな行番号を取得できます()
        {
            Assert.Equal((uint)16384, new CellName(16384, 1).ColumnIndex);
        }

        [Fact]
        public void RowIndexプロパティで行番号を取得できます()
        {
            Assert.Equal((uint)1, CellName.Parse("A1").RowIndex);
            Assert.Equal((uint)2, CellName.Parse("A2").RowIndex);
        }

        [Fact]
        public void 行列番号で生成した場合にRowIndexプロパティで行番号を取得できます()
        {
            Assert.Equal((uint)1, new CellName(1, 1).RowIndex);
            Assert.Equal((uint)2, new CellName(1, 2).RowIndex);
        }

        [Fact]
        public void Parseメソッドで生成した場合にRowIndexプロパティで大きな行番号を取得できます()
        {
            Assert.Equal((uint)1048576, CellName.Parse("A1048576").RowIndex);
        }

        [Fact]
        public void コンストラクタで生成した場合にRowIndexプロパティで大きな行番号を取得できます()
        {
            Assert.Equal((uint)1048576, new CellName(1, 1048576).RowIndex);
        }

        [Fact]
        public void ToStringメソッドはA1形式で文字列を返します()
        {
            Assert.Equal("A1", CellName.Parse("A1").ToString());
            Assert.Equal("B2", CellName.Parse("B2").ToString());
            Assert.Equal("AA1", CellName.Parse("AA1").ToString());
            Assert.Equal("AAA1", CellName.Parse("AAA1").ToString());
            Assert.Equal("XFD1048576", CellName.Parse("XFD1048576").ToString());
        }
    }
}
