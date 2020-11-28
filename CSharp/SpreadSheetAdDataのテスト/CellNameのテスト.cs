using Marimo.SpreadSheetAsData;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework.Constraints;

namespace Marimo.SpreadSheetAdData.Test
{
    [TestFixture]
    public class CellNameのテスト
    {
        [Test]
        public void ParseメソッドでCellNameが生成できます()
        {
            Assert.That(() => CellName.Parse("A1"), Throws.Nothing);
        }

        [Test]
        public void 列番号と行番号を指定してCellNameがせいせいできます()
        {
            Assert.That(() => new CellName(1,1), Throws.Nothing);
        }

        [Test]
        public void 同じセル位置は同一として扱われます()
        {
            Assert.That(CellName.Parse("A1"), Is.EqualTo(CellName.Parse("A1")));
        }

        [Test]
        public void 行が違うと同一ではないとして扱われます()
        {
            Assert.That(CellName.Parse("A1"), Is.Not.EqualTo(CellName.Parse("A2")));
        }

        [Test]
        public void 列が違うと同一ではないとして扱われます()
        {
            Assert.That(CellName.Parse("A1"), Is.Not.EqualTo(CellName.Parse("B1")));
        }

        [Test]
        public void 指定したセルの位置を一意としてハッシュのキーとして使えます()
        {
            var set = new HashSet<CellName>();

            set.Add(CellName.Parse("A1"));
            Assert.That(set.Count, Is.EqualTo(1));
            set.Add(CellName.Parse("A1"));
            Assert.That(set.Count, Is.EqualTo(1));
            set.Add(CellName.Parse("B1"));
            Assert.That(set.Count, Is.EqualTo(2));
            set.Add(CellName.Parse("A2"));
            Assert.That(set.Count, Is.EqualTo(3));
        }

        [Test]
        public void Parseメソッドに無効なセル名を指定するとFormatExceptionを投げます()
        {
            Assert.That(() => CellName.Parse("A"), Throws.InstanceOf<FormatException>());
            Assert.That(() => CellName.Parse("1"), Throws.InstanceOf<FormatException>());
            Assert.That(() => CellName.Parse("A1A1"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Parseメソッドに大きすぎる列名を指定した場合はFormatExceptionを投げます()
        {
            Assert.That(() => CellName.Parse("XFE1"), Throws.InstanceOf<FormatException>());
            Assert.That(() => CellName.Parse("AAAA1"), Throws.InstanceOf<FormatException>());
            Assert.That(() => CellName.Parse("ZZZZ1"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void コンストラクタに大きすぎる列番号を指定した場合はFormatExceptionを投げます()
        {
            Assert.That(() => new CellName(16385, 1), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Parseメソッドに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
        {
            Assert.That(() => CellName.Parse("A1048577"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void コンストラクタに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
        {
            Assert.That(() => new CellName(1, 1048577), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void ColumnNameプロパティで列名を取得できます()
        {
            Assert.That(CellName.Parse("A1").ColumnName, Is.EqualTo("A"));
            Assert.That(CellName.Parse("B1").ColumnName, Is.EqualTo("B"));
        }

        [Test]
        public void 行列番号で生成した場合にColumnNameプロパティで列名を取得できます()
        {
            Assert.That(new CellName(1, 1).ColumnName, Is.EqualTo("A"));
            Assert.That(new CellName(2, 1).ColumnName, Is.EqualTo("B"));
        }

        [Test]
        public void ColumnIndexプロパティで一文字の列番号を取得できます()
        {
            Assert.That(CellName.Parse("A1").ColumnIndex, Is.EqualTo(1));
            Assert.That(CellName.Parse("Z1").ColumnIndex, Is.EqualTo(26));
        }

        [Test]
        public void ColumnIndexプロパティで二文字のの列番号を取得できます()
        {
            Assert.That(CellName.Parse("AA1").ColumnIndex, Is.EqualTo(27));
            Assert.That(CellName.Parse("ZZ1").ColumnIndex, Is.EqualTo(702));
        }

        [Test]
        public void ColumnIndexプロパティで三文字のの列番号を取得できます()
        {
            Assert.That(CellName.Parse("AAA1").ColumnIndex, Is.EqualTo(703));
            Assert.That(CellName.Parse("XFD1").ColumnIndex, Is.EqualTo(16384));
        }

        [Test]
        public void コンストラクタで生成した場合にColumnIndexプロパティで大きな行番号を取得できます()
        {
            Assert.That(new CellName(16384, 1).ColumnIndex, Is.EqualTo(16384));
        }

        [Test]
        public void RowIndexプロパティで行番号を取得できます()
        {
            Assert.That(CellName.Parse("A1").RowIndex, Is.EqualTo(1));
            Assert.That(CellName.Parse("A2").RowIndex, Is.EqualTo(2));
        }

        [Test]
        public void 行列番号で生成した場合にRowIndexプロパティで行番号を取得できます()
        {
            Assert.That(new CellName(1, 1).RowIndex, Is.EqualTo(1));
            Assert.That(new CellName(1, 2).RowIndex, Is.EqualTo(2));
        }

        [Test]
        public void Parseメソッドで生成した場合にRowIndexプロパティで大きな行番号を取得できます()
        {
            Assert.That(CellName.Parse("A1048576").RowIndex, Is.EqualTo(1048576));
        }

        [Test]
        public void コンストラクタで生成した場合にRowIndexプロパティで大きな行番号を取得できます()
        {
            Assert.That(new CellName(1, 1048576).RowIndex, Is.EqualTo(1048576));
        }

        [Test]
        public void ToStringメソッドはA1形式で文字列を返します()
        {
            Assert.That(CellName.Parse("A1").ToString(), Is.EqualTo("A1"));
            Assert.That(CellName.Parse("B2").ToString(), Is.EqualTo("B2"));
            Assert.That(CellName.Parse("AA1").ToString(), Is.EqualTo("AA1"));
            Assert.That(CellName.Parse("AAA1").ToString(), Is.EqualTo("AAA1"));
            Assert.That(CellName.Parse("XFD1048576").ToString(), Is.EqualTo("XFD1048576"));
        }
    }
}
