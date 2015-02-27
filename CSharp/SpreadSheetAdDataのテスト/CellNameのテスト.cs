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
    public class CellNameのテスト : AssertionHelper
    {
        [Test]
        public void ParseメソッドでCellNameが生成できます()
        {
            Expect(() => CellName.Parse("A1"), Throws.Nothing);
        }

        [Test]
        public void 同じセル位置は同一として扱われます()
        {
            Expect(CellName.Parse("A1"), Is.EqualTo(CellName.Parse("A1")));
        }

        [Test]
        public void 行が違うと同一ではないとして扱われます()
        {
            Expect(CellName.Parse("A1"), Is.Not.EqualTo(CellName.Parse("A2")));
        }

        [Test]
        public void 列が違うと同一ではないとして扱われます()
        {
            Expect(CellName.Parse("A1"), Is.Not.EqualTo(CellName.Parse("B1")));
        }

        [Test]
        public void 指定したセルの位置を一意としてハッシュのキーとして使えます()
        {
            var set = new HashSet<CellName>();

            set.Add(CellName.Parse("A1"));
            Expect(set.Count, Is.EqualTo(1));
            set.Add(CellName.Parse("A1"));
            Expect(set.Count, Is.EqualTo(1));
            set.Add(CellName.Parse("B1"));
            Expect(set.Count, Is.EqualTo(2));
            set.Add(CellName.Parse("A2"));
            Expect(set.Count, Is.EqualTo(3));
        }

        [Test]
        public void Parseメソッドに無効なセル名を指定するとFormatExceptionを投げます()
        {
            Expect(() => CellName.Parse("A"), Throws.InstanceOf<FormatException>());
            Expect(() => CellName.Parse("1"), Throws.InstanceOf<FormatException>());
            Expect(() => CellName.Parse("A1A1"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Parseメソッドに大きすぎる列名を指定した場合はFormatExceptionを投げます()
        {
            Expect(() => CellName.Parse("XFE1"), Throws.InstanceOf<FormatException>());
            Expect(() => CellName.Parse("AAAA1"), Throws.InstanceOf<FormatException>());
            Expect(() => CellName.Parse("ZZZZ1"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void Parseメソッドに大きすぎる行番号を指定した場合はFormatExceptionを投げます()
        {
            Expect(() => CellName.Parse("A1048577"), Throws.InstanceOf<FormatException>());
        }

        [Test]
        public void ColumnNameプロパティで列名を取得できます()
        {
            Expect(CellName.Parse("A1").ColumnName, Is.EqualTo("A"));
            Expect(CellName.Parse("B1").ColumnName, Is.EqualTo("B"));
        }

        [Test]
        public void ColumnIndexプロパティで一文字の列番号を取得できます()
        {
            Expect(CellName.Parse("A1").ColumnIndex, Is.EqualTo(1));
            Expect(CellName.Parse("Z1").ColumnIndex, Is.EqualTo(26));
        }

        [Test]
        public void ColumnIndexプロパティで二文字のの列番号を取得できます()
        {
            Expect(CellName.Parse("AA1").ColumnIndex, Is.EqualTo(27));
            Expect(CellName.Parse("ZZ1").ColumnIndex, Is.EqualTo(702));
        }

        [Test]
        public void ColumnIndexプロパティで三文字のの列番号を取得できます()
        {
            Expect(CellName.Parse("AAA1").ColumnIndex, Is.EqualTo(703));
            Expect(CellName.Parse("XFD1").ColumnIndex, Is.EqualTo(16384));
        }

        [Test]
        public void RowIndexプロパティで行番号を取得できます()
        {
            Expect(CellName.Parse("A1").RowIndex, Is.EqualTo(1));
            Expect(CellName.Parse("A2").RowIndex, Is.EqualTo(2));
        }

        [Test]
        public void RowIndexプロパティで大きな行番号を取得できます()
        {
            Expect(CellName.Parse("A1048576").RowIndex, Is.EqualTo(1048576));
        }

        [Test]
        public void ToStringメソッドはA1形式で文字列を返します()
        {
            Expect(CellName.Parse("A1").ToString(), Is.EqualTo("A1"));
            Expect(CellName.Parse("B2").ToString(), Is.EqualTo("B2"));
            Expect(CellName.Parse("XFD1048576").ToString(), Is.EqualTo("XFD1048576"));
        }
    }
}
